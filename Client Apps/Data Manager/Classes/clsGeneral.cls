VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private datData As clsDataAccess

Private asViewTables() As String



Public Function DictionaryCompareStrings(psString1 As Variant, psString2 As Variant) As Integer

  ' String comparison function which equates to SQL Server's dictionary
  ' sort order.
  ' Returns -1 if psString1 < psString2
  ' Returns  0 if psString1 = psString2
  ' Returns  1 if psString1 > psString2
  Dim iLoop As Integer
  Dim iResult As Integer
  Dim iSubResult As Integer
  Dim iMinLength As Integer
  Dim iFirstAccent As Integer
  Dim sString1 As String
  Dim sString2 As String
  Dim sChar1 As String
  Dim sChar2 As String
  Dim iGroup1 As Integer
  Dim iGroup2 As Integer
  Dim sRealChar1 As String
  Dim sRealChar2 As String
  
  If IsNull(psString1) Or IsNull(psString2) Then
    If IsNull(psString1) And IsNull(psString2) Then
      iResult = 0
    Else
      iResult = IIf(IsNull(psString1), -1, 1)
    End If
  Else
    ' JPD 28/6/01 Use RTRIM to avoid 'abc' and 'abc ' being thought of as different
    ' (SQL doesn't think so, so why should we).
    sString1 = RTrim(CStr(psString1))
    sString2 = RTrim(CStr(psString2))
    
    iFirstAccent = 0
    iMinLength = IIf(Len(sString1) < Len(sString2), Len(sString1), Len(sString2))
    
    For iLoop = 1 To iMinLength
      sChar1 = Mid(sString1, iLoop, 1)
      sChar2 = Mid(sString2, iLoop, 1)
      
      iGroup1 = DictionaryOrderGroup(sChar1)
      iGroup2 = DictionaryOrderGroup(sChar2)
      
      If iGroup1 < iGroup2 Then
        iResult = -1
      ElseIf iGroup1 > iGroup2 Then
        iResult = 1
      Else
        If (iGroup1 = 4) Or (iGroup1 = 5) Then
          iResult = 0
        ElseIf iGroup1 = 3 Then
          sRealChar1 = DictionaryRealLetter(sChar1)
          sRealChar2 = DictionaryRealLetter(sChar2)
          
          iSubResult = StrComp(sRealChar1, sRealChar2, vbBinaryCompare)
          
          'JPD 20040326 Fault 8335
          If (iSubResult = 0) _
            And ((sChar1 = "ß") Or (sChar2 = "ß")) Then
            
            iSubResult = StrComp(UCase(sChar1), UCase(sChar2), vbBinaryCompare)
          End If
          
          If iSubResult = 0 Then
            If iFirstAccent = 0 Then
              ' JPD 22/06/01 Compare the ucase of the characters.
              ' This fixes the bug which sees 'ABC' and 'Abc' as having
              ' an accent on the 'b'.
              iFirstAccent = StrComp(UCase(sChar1), UCase(sChar2), vbBinaryCompare)
            End If
          Else
            iResult = iSubResult
          End If
        Else
          iResult = StrComp(sChar1, sChar2, vbBinaryCompare)
        End If
      End If
      
      If iResult <> 0 Then
        Exit For
      End If
    Next iLoop
    
    If iResult = 0 Then
      If Len(sString1) < Len(sString2) Then
        iResult = -1
      ElseIf Len(sString1) > Len(sString2) Then
        iResult = 1
      End If
    End If
  
    If (iResult = 0) And (iFirstAccent <> 0) Then
      iResult = iFirstAccent
    End If
  End If
  
  DictionaryCompareStrings = iResult
  
End Function

Private Function DictionaryOrderGroup(psString As String) As Integer

  ' Used by the DictionaryCompareStrings function.
  ' Returns the order group for the given character.
  Dim iResult As Integer
  Dim iAsciiCode As Integer
  
  iResult = 0
  iAsciiCode = Asc(psString)
  
  If ((iAsciiCode >= 32) And (iAsciiCode <= 47)) Or _
    ((iAsciiCode >= 58) And (iAsciiCode <= 64)) Or _
    ((iAsciiCode >= 91) And (iAsciiCode <= 96)) Or _
    ((iAsciiCode >= 123) And (iAsciiCode <= 191)) Or _
    (iAsciiCode = 215) Or _
    (iAsciiCode = 247) Then
    ' <Space>!"#$%&'()*+,-./
    ' :;<=>?@
    ' [\]^_`
    ' {|}~‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ<blank>¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿
    ' ×
    ' ÷
    iResult = 1
  ElseIf (iAsciiCode >= 48) And (iAsciiCode <= 57) Then
    ' 0 to 9
    iResult = 2
  ElseIf ((iAsciiCode >= 65) And (iAsciiCode <= 90)) Or _
    ((iAsciiCode >= 97) And (iAsciiCode <= 122)) Or _
    ((iAsciiCode >= 192) And (iAsciiCode <= 207)) Or _
    ((iAsciiCode >= 209) And (iAsciiCode <= 214)) Or _
    ((iAsciiCode >= 216) And (iAsciiCode <= 221)) Or _
    ((iAsciiCode >= 223) And (iAsciiCode <= 239)) Or _
    ((iAsciiCode >= 241) And (iAsciiCode <= 246)) Or _
    ((iAsciiCode >= 248) And (iAsciiCode <= 253)) Or _
    (iAsciiCode = 255) Then
    ' A to Z
    ' a to z
    ' ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏ
    ' ÑÒÓÔÕÖ
    ' ØÙÚÛÜÝ
    ' ßàáâãäåæçèéêëìíîï
    ' ñòóôõö
    ' øùúûüý
    ' ÿ
    iResult = 3
  ElseIf (iAsciiCode = 208) Or (iAsciiCode = 240) Then
    ' Ðð
    iResult = 4
  ElseIf (iAsciiCode = 222) Or (iAsciiCode = 254) Then
    ' Þþ
    iResult = 5
  End If
  
  DictionaryOrderGroup = iResult

End Function

Private Function DictionaryRealLetter(ByVal psChar As String) As String

  ' Used by the DictionaryCompareStrings function.
  ' Convert lowercase and accented characters into their
  ' uppercase non-accented versions.
  Dim iAsciiCode As Integer
    
  psChar = UCase(psChar)
  iAsciiCode = Asc(psChar)
  
  If ((iAsciiCode >= 192) And (iAsciiCode <= 198)) Then
    ' ÀÁÂÃÄÅÆ
    psChar = "A"
  ElseIf (iAsciiCode = 199) Then
    ' Ç
    psChar = "C"
  ElseIf ((iAsciiCode >= 200) And (iAsciiCode <= 203)) Then
    ' ÈÉÊË
    psChar = "E"
  ElseIf ((iAsciiCode >= 204) And (iAsciiCode <= 207)) Then
    ' ÌÍÎÏ
    psChar = "I"
  ElseIf (iAsciiCode = 209) Then
    ' Ñ
    psChar = "N"
  ElseIf ((iAsciiCode >= 210) And (iAsciiCode <= 214)) Or _
    (iAsciiCode = 216) Then
    ' ÒÓÔÕÖØ
    psChar = "O"
  ElseIf ((iAsciiCode >= 217) And (iAsciiCode <= 220)) Then
    ' ÙÚÛÜ
    psChar = "U"
  ElseIf (iAsciiCode = 221) Or (iAsciiCode = 159) Then
    ' Ý
    psChar = "Y"
  ElseIf (iAsciiCode = 223) Then
    ' ß
    psChar = "S"
  End If
  
  DictionaryRealLetter = psChar

End Function

Public Function FilteredIDs(plngExprID As Long, psIDSQL As String) As Boolean

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.FilteredIDs(plngExprID,psIDSQL)", Array(plngExprID, psIDSQL)
  
  ' Return a string describing the record IDs from the given table
  ' that satisfy the given criteria.
  Dim fOK As Boolean
  Dim objExpr As clsExprExpression

  Screen.MousePointer = vbHourglass

  Set objExpr = New clsExprExpression
  With objExpr
    ' Initialise the filter expression object.
    fOK = .Initialise(0, plngExprID, giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
    
    If fOK Then
      fOK = objExpr.RuntimeFilterCode(psIDSQL, True, False)
    End If
    
  End With
  Set objExpr = Nothing

  Screen.MousePointer = vbDefault
  
  FilteredIDs = fOK

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function


Public Function GetValueForRecordIndependantCalc(lngExprID As Long) As Variant

  Dim objExpr As clsExprExpression
  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim fOK As Boolean
  Dim lngViews() As Long
  
  ReDim lngViews(0)
  
  
  On Local Error GoTo LocalErr
  
  GetValueForRecordIndependantCalc = vbNullString
  
  Set objExpr = New clsExprExpression
  With objExpr
    ' Initialise the filter expression object.
    fOK = .Initialise(0, lngExprID, giEXPR_RECORDINDEPENDANTCALC, giEXPRVALUE_UNDEFINED)
    
    If fOK Then
      fOK = objExpr.RuntimeCalculationCode(lngViews, strSQL, True)
    End If
    
    If fOK Then
      Set rsTemp = GetReadOnlyRecords("SELECT " & strSQL)
      If Not rsTemp.BOF And Not rsTemp.EOF Then
        GetValueForRecordIndependantCalc = rsTemp.Fields(0).Value
      End If
    End If

  End With
  Set objExpr = Nothing


Exit Function

LocalErr:

End Function


Public Function GetColumnNames(plngTableID As Long, ByVal IncludeIDs As Boolean) As ADODB.Recordset

  ' Return a recordset of the the columns in the given table.
  Dim sSQL As String
  Dim rsColumns As ADODB.Recordset

  sSQL = "SELECT ColumnName, ColumnID" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & Trim(Str(plngTableID)) & IIf(Not IncludeIDs, " AND columntype <> 3", "") & _
    " ORDER BY ColumnName;"
  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetColumnNames = rsColumns

End Function

Public Function GetOrderDefinition(plngOrderID As Long) As ADODB.Recordset

  ' Return a recordset of the order items (both Find Window and Sort Order columns)
  ' for the given order.
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
    
  sSQL = "EXEC dbo.sp_ASRGetOrderDefinition " & plngOrderID
  Set rsInfo = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetOrderDefinition = rsInfo

End Function

Public Sub ClearConnection()

  If Not gADOCon Is Nothing Then
    If gADOCon.State = adStateOpen Then
      gADOCon.Close
    End If
    Set gADOCon = Nothing
  End If

End Sub

Public Function Connect(sConnect As String, sMsg As String, sUser As String, lngTimeOut As Long, ByRef bDatabaseLocked As Boolean) As Boolean

  Dim rsUsers As ADODB.Recordset
   
  Dim sSQL As String
  
  On Error GoTo Err_Trap
  
  Set gADOCon = New ADODB.Connection
  With gADOCon
    .ConnectionString = sConnect
    .CommandTimeout = lngTimeOut
    .ConnectionTimeout = lngTimeOut
    .CursorLocation = adUseServer
    .Mode = adModeReadWrite
    .Open
  End With
  
  ' Check if the user is a SQL Server System Administrator.
  ' We do not allow these users to login to the data manager.
  If Not ASRDEVELOPMENT Then
    
    sSQL = "SELECT IS_SRVROLEMEMBER('sysadmin') AS sysadmin" & _
      ",IS_SRVROLEMEMBER('securityadmin') AS securityadmin" & _
      ",IS_SRVROLEMEMBER('serverAdmin') AS serverAdmin" & _
      ",IS_SRVROLEMEMBER('setupAdmin') AS setupAdmin" & _
      ",IS_SRVROLEMEMBER('processAdmin') AS processAdmin" & _
      ",IS_SRVROLEMEMBER('diskAdmin') AS diskAdmin" & _
      ",IS_SRVROLEMEMBER('dbCreator') AS dbCreator"

    Set rsUsers = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsUsers
      If Not (.BOF And .EOF) Then
        If (!sysAdmin = 1) Or _
          (!securityAdmin = 1) Or _
          (!serverAdmin = 1) Or _
          (!setupAdmin = 1) Or _
          (!processAdmin = 1) Or _
          (!diskAdmin = 1) Or _
          (!dbCreator = 1) Then
  
            sMsg = "Unable to login to the Data Manager." & vbCrLf & _
              "Users assigned to fixed SQL Server roles cannot use the Data Manager."
            GoTo Exit_Fail
        
        End If
      End If
      .Close
    End With
  End If

  sSQL = "sp_ASRLockCheck"
  ' AE20080425 Fault #13091
  'Set rsUsers = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set rsUsers = New ADODB.Recordset
  rsUsers.Open "sp_ASRLockCheck", gADOCon, adOpenForwardOnly, adLockReadOnly
    
  With rsUsers
    If Not .BOF And Not .EOF Then
      'Ignore read only locks...
      
      'AE20071213 #S000738
      Select Case rsUsers!Priority
      Case 2 ' Manual lock - Display user message
        Dim sSystemSettingMsg As String
        
        sMsg = "Unable to login to " & gsDatabaseName & " as the database has been locked."
        
        sSystemSettingMsg = GetSystemSetting("messaging", "lockmessage", vbNullString)
        If sSystemSettingMsg <> vbNullString Then
          sMsg = sMsg & vbCrLf & vbCrLf & sSystemSettingMsg
        End If
        
        sMsg = sMsg & vbCrLf & vbCrLf & _
              "User :  " & !userName & vbCrLf & _
              "Date/Time :  " & !Lock_Time & vbCrLf & _
              "Machine :  " & !HostName & vbCrLf & _
              "Type :  " & !Description
        
        Connect = False
        bDatabaseLocked = True
        .Close
        Exit Function
      Case 3 ' ReadWrite
        ' Do Nothing
      Case Else ' System is locked for a save or somthing
        sMsg = "Unable to login to " & gsDatabaseName & " as the database has been locked." & vbCrLf & vbCrLf & _
               "User :  " & !userName & vbCrLf & _
               "Date/Time :  " & !Lock_Time & vbCrLf & _
               "Machine :  " & !HostName & vbCrLf & _
               "Type :  " & !Description
        Connect = False
        bDatabaseLocked = True
        .Close
        Exit Function
      End Select

'      If rsUsers!Priority <> 3 Then
'        sMsg = "Unable to login to " & gsDatabaseName & " as the database has been locked." & vbCrLf & vbCrLf & _
'               "User :  " & !UserName & vbCrLf & _
'               "Date/Time :  " & !Lock_Time & vbCrLf & _
'               "Machine :  " & !HostName & vbCrLf & _
'               "Type :  " & !Description
'        Connect = False
'        bDatabaseLocked = True
'        .Close
'        Exit Function
'      End If
    End If
    .Close
  End With


Exit_Ok:
  Connect = True
  gsConnectString = sConnect
  If rsUsers.State = adStateOpen Then
    rsUsers.Close
  End If
  Set rsUsers = Nothing
  
  Exit Function
    
Exit_Fail:
    Connect = False
    If rsUsers.State = adStateOpen Then
      rsUsers.Close
    End If
    Set rsUsers = Nothing
    Exit Function

Err_Trap:
    Connect = False

End Function

Public Function CheckLicence() As Boolean

  Dim sMsg As String
  Dim lActualCount As Long
  Dim lngCurrentHeadcount As Long
  Dim dToday As Date
  
  On Error GoTo Err_Trap
  
  CheckLicence = False
  dToday = DateValue(Now)
    
  ' Expiry date checks
  If gobjLicence.ExpiryDate <> "00:00:00" Then
    If (dToday > gobjLicence.ExpiryDate) Then
      sMsg = "Your licence to use this product has expired." & vbNewLine & vbNewLine & _
            "Please contact OpenHR Customer Services on 08451 609 999 as soon as possible."
      GoTo Exit_Fail
    End If
            
    If (dToday > DateAdd("d", -7, gobjLicence.ExpiryDate)) Then
      sMsg = "Your licence to use this product will expire on " & gobjLicence.ExpiryDate & "." & vbNewLine & vbNewLine & _
            "Please contact OpenHR Customer Services on 08451 609 999 as soon as possible."
      COAMsgBox sMsg, vbInformation
    End If
  End If
     
     
  ' Concurrency checks
  If gobjLicence.LicenceType = Concurrency Or gobjLicence.LicenceType = DMIConcurrencyAndHeadcount Or gobjLicence.LicenceType = DMIConcurrencyAndP14 Then
    lActualCount = GetCurrentUsersCountInApp
    
    If lActualCount > gobjLicence.DATUsers Then
      sMsg = "The maximum number of licenced users are currently" & vbNewLine & "logged into OpenHR - Please try again later." & vbNewLine & vbNewLine & _
             "If you wish to increase the number of licenced users, please" & vbNewLine & "contact OpenHR Customer Services on 08451 609 999 as soon as possible."
      GoTo Exit_Fail
    End If
  End If


  ' Headcount checks
  If gobjLicence.Headcount > 0 Then
    Select Case gobjLicence.LicenceType
      Case LicenceType.Headcount, LicenceType.DMIConcurrencyAndHeadcount
        lngCurrentHeadcount = GetSystemSetting("Headcount", "current", 0)
        
      Case LicenceType.P14Headcount, LicenceType.DMIConcurrencyAndP14
       lngCurrentHeadcount = GetSystemSetting("Headcount", "P14", 0)

    End Select
   
    If lngCurrentHeadcount >= gobjLicence.Headcount Then
      sMsg = "You have reached or exceeded the headcount limit set within the terms of your licence agreement." & vbNewLine & vbNewLine & _
                            "You are no longer able to add new employee records, but you may access the system for other purposes." & vbNewLine & vbNewLine & _
                            "Please contact OpenHR Customer Services on 08451 609 999 as soon as possible to increase the licence headcount number."
      COAMsgBox sMsg, vbCritical
    
    ElseIf lngCurrentHeadcount >= gobjLicence.Headcount * 0.95 Then
      
      If DisplayWarningToUser(gsUserName, Headcount95Percent, 7) Then
        sMsg = "You are currently within 95% (" & lngCurrentHeadcount & " of " & gobjLicence.Headcount & " employees) of reaching the headcount limit set within the terms of your licence agreement." & vbNewLine & vbNewLine & _
                              "Once this limit is reached, you will no longer be able to add new employee records to the system." & vbNewLine & vbNewLine & _
                              "If you wish to increase the headcount number, please contact" & vbNewLine & vbNewLine & _
                              "OpenHR Customer Services on 08451 609 999 as soon as possible."
        COAMsgBox sMsg, vbInformation
      End If
    
    End If
  End If
    

Exit_Ok:
  CheckLicence = True
  Exit Function
    
Exit_Fail:
    COAMsgBox sMsg, vbCritical
    CheckLicence = False
    Exit Function

Err_Trap:
    CheckLicence = False

End Function


Public Property Let UserNameForSQL(ByVal sUsername As String)
  gsUserName = sUsername
  gsUserGroup = GetUserDetails
  gfCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr

End Property

Public Property Get UserNameForSQL() As String

  If InStrB(1, gsUserName, "'") Then
    UserNameForSQL = Replace(gsUserName, "'", "''")
  Else
    UserNameForSQL = gsUserName
  End If

End Property

Public Function GetScreenTables() As ADODB.Recordset

  Dim sSQL As String
  Dim rsTables As ADODB.Recordset
  
  sSQL = "SELECT DISTINCT ASRSysTables.TableID , ASRSysTables.TableName " & _
      "FROM ASRSysTables INNER JOIN ASRSysScreens ON ASRSysTables.TableID = " & _
      "ASRSysScreens.TableID Where ASRSysTables.TableType = 1 Order By ASRSysTables.TableName Desc"
  
  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetScreenTables = rsTables
    
End Function

Public Function GetPrimaryTableMenu() As Variant

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetPrimaryTableMenu()"

  ' Return an array of information that can be used to format the main Database menu for the current user.
  ' The recordset contains a row for each primary table in the OpenHR database.
  ' For each primary table the following information is given :
  '  tableID     ID of the primary table
  '  tableName   Name of the primary table
  '  tableScreenCount  Number of screens associated with the primary table
  '  tableScreenID   If the current user has SELECT permission on the primary table, and the primary table has only one screen associated with it,
  '        and (the current user does not have SELECT permission on any views on the primary table, or there are no screens associated with the view)
  '        then the ID of the one screen associated with the primary table is returned
  '  tableReadable   True if the current user has SELECT permission on the primary table
  '  tableViewCount    Number of views on the primary table that the current user has SELECT permission on
  '  viewID      If (the current user does not have SELECT permission on the primary table, or the primary table has no screens associated with it),
  '        and (the current user has SELECT permission on only one view on the primary table, and there are is only one screen associated with the view)
  '        then the ID of the view is returned
  '  viewName    If (the current user does not have SELECT permission on the primary table, or the primary table has no screens associated with it),
  '        and (the current user has SELECT permission on only one view on the primary table, and there are is only one screen associated with the view)
  '        then the name of the view is returned
  '  viewScreenCount Number of screens associated with the views
  '  viewScreenID    If (the current user does not have SELECT permission on the primary table, or the primary table has no screens associated with it),
  '        and (the current user has SELECT permission on only one view on the primary table, and there are is only one screen associated with the view)
  '        then the name of the screen associated with the view is returned
  '  tableScreenPictureID - ID of the table screen's icon
  '  viewScreenPictureID - ID of the view screen's icon
  '
  ' If the recordset for a primary table has a non-zero value in the tableScreenID field then the table just requires a tool on the Database menu that calls up the given screen.
  ' Else, if the recordset for a primary table has a non-zero value in the viewID field then the table just requires a tool on the Database menu that calls up the given screen for the permitted view.
  ' Else, if the recordset for a primary table has a non-zero value in the viewScreenCount field OR (the tableReadable value is 1 AND the tableScreenCount is greater than 1) then
  '  the table requires a tool on the Database menu that calls up a sub-band of the collection of views/screens available for the primary table.
  ' Else, the primary table should not appear on the menu.
  Dim iNextIndex As Integer
  Dim iTotalViewScreenCount As Integer
  Dim sSQL As String
  Dim sViewList As String
  Dim rsViews As ADODB.Recordset
  Dim rsViewScreen As ADODB.Recordset
  Dim rsTables As ADODB.Recordset
  Dim rsTableScreen As ADODB.Recordset
  Dim avTableInfo() As Variant
  Dim objTableView As CTablePrivilege

  ReDim avTableInfo(12, 0)
  ' Index 1 = table ID
  ' Index 2 = table name
  ' Index 3 = table screen count
  ' Index 4 = table screen ID
  ' Index 5 = table readable
  ' Index 6 = table view count
  ' Index 7 = view ID
  ' Index 8 = view name
  ' Index 9 = view screen count
  ' Index 10 = view screen ID
  ' Index 11 = table screen picture ID
  ' Index 12 = view screen picture ID
  
  ' Get a recordset of the primary tables in the database.
  sSQL = "SELECT ASRSysTables.tableID," & _
    " ASRSysTables.tableName," & _
    " COUNT(DISTINCT ASRSysScreens.name) AS tableScreenCount" & _
    " FROM ASRSysTables" & _
    " INNER JOIN ASRSysScreens" & _
    " ON ASRSysTables.tableID = ASRSysScreens.tableID" & _
    " AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))" & _
    " AND ((ASRSysScreens.quickEntry IS null) OR (ASRSysScreens.quickEntry = 0))" & _
    " GROUP BY ASRSysTables.tableID," & _
    " ASRSysTables.tableName," & _
    " ASRSysTables.tableType" & _
    " HAVING ASRSysTables.tableType = 1" & _
    " ORDER BY ASRSysTables.tableName DESC"
  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  Do While Not rsTables.EOF
    ' Initialise an entry into our array of table info for each primary table.
    iNextIndex = UBound(avTableInfo, 2) + 1
    ReDim Preserve avTableInfo(12, iNextIndex)
    
    avTableInfo(1, iNextIndex) = rsTables!TableID
    avTableInfo(2, iNextIndex) = rsTables!TableName
    avTableInfo(3, iNextIndex) = rsTables!tableScreenCount
    avTableInfo(4, iNextIndex) = 0
    avTableInfo(5, iNextIndex) = gcoTablePrivileges.Item(rsTables!TableName).AllowSelect
    avTableInfo(6, iNextIndex) = 0
    avTableInfo(7, iNextIndex) = 0
    avTableInfo(8, iNextIndex) = ""
    avTableInfo(9, iNextIndex) = 0
    avTableInfo(10, iNextIndex) = 0
    avTableInfo(11, iNextIndex) = 0
    avTableInfo(12, iNextIndex) = 0
      
    iTotalViewScreenCount = 0
    
    ' Create a list of the current user's permitted views on the current table.
    sViewList = "0"
    For Each objTableView In gcoTablePrivileges.Collection
      If Not (objTableView.IsTable) And _
        (objTableView.TableID = rsTables!TableID) And _
        (objTableView.AllowSelect) Then
        
        sViewList = sViewList & ", " & Trim(Str(objTableView.ViewID))
      End If
    Next objTableView
    Set objTableView = Nothing

    ' Get view information for the current table.
    sSQL = "SELECT ASRSysViews.viewName, COUNT (ASRSysViewScreens.ScreenID) AS viewScreenCount" & _
      " FROM ASRSysViews" & _
      " LEFT OUTER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID" & _
      " GROUP BY ASRSysViews.viewName, ASRSysViews.viewID" & _
      " HAVING ASRSysViews.viewID IN (" & sViewList & ")"
    Set rsViews = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Do While Not rsViews.EOF
      If rsViews!viewScreenCount > 0 Then
        avTableInfo(6, iNextIndex) = avTableInfo(6, iNextIndex) + 1
        iTotalViewScreenCount = iTotalViewScreenCount + rsViews!viewScreenCount
      End If
      
      rsViews.MoveNext
    Loop
    rsViews.Close
    Set rsViews = Nothing
    
    ' Get view screen info if required.
    If avTableInfo(6, iNextIndex) > 0 Then
      avTableInfo(9, iNextIndex) = iTotalViewScreenCount
      
      If (iTotalViewScreenCount = 1) And _
        ((avTableInfo(5, iNextIndex) = False) Or (avTableInfo(3, iNextIndex) = 0)) Then
        
        sSQL = "SELECT ASRSysViews.viewID, ASRSysViews.viewName, ASRSysViewScreens.screenID, ASRSysScreens.pictureID" & _
          " FROM ASRSysViews" & _
          " INNER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID" & _
          " INNER JOIN ASRSysScreens ON ASRSysViewScreens.screenID = ASRSysScreens.screenID" & _
          " WHERE ASRSysViews.viewid IN (" & sViewList & ")"
        Set rsViewScreen = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
        If Not (rsViewScreen.EOF And rsViewScreen.BOF) Then
          avTableInfo(7, iNextIndex) = rsViewScreen!ViewID
          avTableInfo(8, iNextIndex) = rsViewScreen!ViewName
          avTableInfo(10, iNextIndex) = rsViewScreen!ScreenID
          avTableInfo(12, iNextIndex) = rsViewScreen!PictureID
        End If
        rsViewScreen.Close
        Set rsViewScreen = Nothing
      End If
    End If
        
    If (avTableInfo(3, iNextIndex) = 1) And _
      (iTotalViewScreenCount = 0) And _
      (avTableInfo(5, iNextIndex) = True) Then
      sSQL = "SELECT ASRSysScreens.screenID," & _
        " ASRSysScreens.pictureID" & _
        " FROM ASRSysScreens" & _
        " WHERE ASRSysScreens.tableID = " & Trim(Str(rsTables!TableID)) & _
        "   AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))"

      Set rsTableScreen = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      If Not (rsTableScreen.EOF And rsTableScreen.BOF) Then
        avTableInfo(4, iNextIndex) = rsTableScreen!ScreenID
        avTableInfo(11, iNextIndex) = rsTableScreen!PictureID
      End If
      rsTableScreen.Close
      Set rsTableScreen = Nothing
    End If
            
    rsTables.MoveNext
  Loop
  rsTables.Close
  Set rsTables = Nothing
  GetPrimaryTableMenu = avTableInfo
    
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

Public Function GetScreenScreens(plngTableID As Long) As ADODB.Recordset

  ' Get the complete set of screens for the given table.
  Dim sSQL As String
  Dim rsScreens As ADODB.Recordset
  
  'JPD 20041215 Fault 9637
  sSQL = "SELECT screenID, name, pictureID" & _
    " FROM ASRSysScreens" & _
    " WHERE tableID = " & Trim(Str(plngTableID)) & _
    "   AND isnull(SSIntranet, 0) = 0" & _
    " ORDER BY name DESC"
    
  Set rsScreens = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
  Set GetScreenScreens = rsScreens

End Function

'MH20020808 These Functions are not called from anywhere.
'Public Function GetScreenViews(plngTableID As Long) As ADODB.Recordset
'
'  ' Return a recordset of the views on the given table.
'  Dim sSQL As String
'  Dim rsViews As ADODB.Recordset
'
'  sSQL = "SELECT viewAlternativeName AS name, viewName, viewID" & _
'    " FROM ASRSysViews" & _
'    " WHERE viewTableID = " & plngTableID & _
'    " ORDER BY name DESC"
'
'  Set rsViews = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
'  Set GetScreenViews = rsViews
'
'  Set rsViews = Nothing
'
'End Function
'
'Public Function GetScreenViews2(lScreenID As Long) As ADODB.Recordset
'
'  Dim sSQL As String
'  Dim rsViews As ADODB.Recordset
'
'  sSQL = "select asrsysviews.viewid,asrsysviews.viewname,asrsysviews.viewalternativename as Name,asrsysviewscreens.screenid,asrsysviewscreens.viewid "
'  sSQL = sSQL & "From asrsysviews, asrsysviewscreens "
'  sSQL = sSQL & "Where asrsysviews.ViewID = asrsysviewscreens.ViewID "
'  sSQL = sSQL & "and asrsysviewscreens.screenid = " & lScreenID
'  sSQL = sSQL & " ORDER BY Name DESC"
'
'  Set rsViews = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
'  Set GetScreenViews2 = rsViews
'
'End Function
'
'Public Function GetViewScreens(lViewID As Long) As ADODB.Recordset
'
'  Dim sSQL As String
'  Dim rsScreens As ADODB.Recordset
'
'  sSQL = "SELECT ASRSysScreens.ScreenID, ASRSysScreens.Name, ASRSysScreens.PictureID " & _
'      "FROM ASRSysScreens INNER JOIN ASRSysViewScreens ON ASRSysScreens.ScreenID = " & _
'      "ASRSysViewScreens.ScreenID Where ASRSysViewScreens.ViewID = " & lViewID
'
'  '#rh 21/7/99
'  sSQL = sSQL & " ORDER BY AsrSysScreens.Name DESC"
'
'  Set rsScreens = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
'  Set GetViewScreens = rsScreens
'
'End Function

Public Function GetAllTables() As ADODB.Recordset

  ' Return a recordset of the user-defined tables.
  Dim sSQL As String
  Dim rsTables As ADODB.Recordset
    
  sSQL = "SELECT tableID, tableName, tableType, defaultOrderID, recordDescExprID, isremoteview" & _
    " FROM ASRSysTables"
    
  Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetAllTables = rsTables

End Function

Public Sub GetColumnPermissions(ByRef pobjColumnPrivileges As CColumnPrivileges)
' JPD20020814 New child views - no longer required
'''  Dim fTopLevelTable As Boolean
'''  Dim fSelectPermission As Boolean
'''  Dim fUpdatePermission As Boolean
'''  Dim sSQL As String
'''  Dim rsColumns As ADODB.Recordset
'''
'''  fTopLevelTable = (gcoTablePrivileges.Item(pobjColumnPrivileges.Tag).IsTable And _
'''    (gcoTablePrivileges.Item(pobjColumnPrivileges.Tag).TableType <> tabChild))
'''
'''  If (Len(gcoTablePrivileges.Item(pobjColumnPrivileges.Tag).RealSource) = 0) Then
'''    sSQL = "exec sp_ASRQuickTableColumnPermissions " & _
'''      "'" & gsUserGroup & "'," & _
'''      "'" & gcoTablePrivileges.Item(pobjColumnPrivileges.Tag).TableName & "'"
'''  Else
'''    sSQL = "exec sp_ASRQuickTableColumnPermissions " & _
'''      "'" & gsUserGroup & "'," & _
'''      "'" & gcoTablePrivileges.Item(pobjColumnPrivileges.Tag).RealSource & "'"
'''  End If
'''
'''  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'''
'''  Do While Not rsColumns.EOF
'''    fSelectPermission = False
'''    fUpdatePermission = False
'''
'''    If rsColumns!Action = 193 Then
'''      fSelectPermission = rsColumns!Permission
'''    End If
'''
'''    If rsColumns!Action = 197 Then
'''      fUpdatePermission = rsColumns!Permission
'''    End If
'''
'''    If Not pobjColumnPrivileges.IsValid(rsColumns!ColumnName) Then
'''      pobjColumnPrivileges.Add fSelectPermission, _
'''        fUpdatePermission, _
'''        rsColumns!ColumnName, _
'''        rsColumns!ColumnType, _
'''        rsColumns!DataType, _
'''        rsColumns!ColumnID, _
'''        (rsColumns!UniqueCheck <> 0)
'''    Else
'''     If rsColumns!Action = 193 Then
'''       pobjColumnPrivileges.Item(rsColumns!ColumnName).AllowSelect = fSelectPermission
'''     End If
'''
'''     If rsColumns!Action = 197 Then
'''       pobjColumnPrivileges.Item(rsColumns!ColumnName).AllowUpdate = fUpdatePermission
'''     End If
'''    End If
'''
'''    rsColumns.MoveNext
'''  Loop
'''  rsColumns.Close
'''  Set rsColumns = Nothing
'''
End Sub



''Public Function GetTablePrivileges(bFirstRun As Boolean) As ADODB.Recordset
''  Dim sSQL As String
''  Dim rsPriv As ADODB.Recordset
''
''  If bFirstRun Then
''    sSQL = "exec sp_ASRTablePrivileges '" & gsUserGroup & "'"
''  Else
''    sSQL = "exec sp_ASRGetTablePrivileges '" & gsUserGroup & "'"
''  End If
''
''  Set rsPriv = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
''  Set GetTablePrivileges = rsPriv
''
''End Function

Public Function GetAllViews() As ADODB.Recordset

  Dim sSQL As String
  Dim rsViews As ADODB.Recordset
    
  sSQL = "SELECT ASRSysViews.viewID, ASRSysViews.viewName, ASRSysTables.tableID," & _
    " ASRSysTables.tableName, ASRSysTables.tableType, ASRSysTables.defaultOrderID, ASRSysTables.recordDescExprID" & _
    " FROM ASRSysViews" & _
    " INNER JOIN ASRSysTables ON ASRSysViews.viewTableID = ASRSysTables.tableID"
  Set rsViews = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetAllViews = rsViews

End Function

Public Function ViewRequiresLocalCursor(plngViewID As Long) As Boolean
  ' Return TRUE if the given view requires a LOCAL cursor.
  ' ie. if it uses a TABLE type UDF in its filter definition.
  ' ie. If it uses a Hierarchy function.
  On Error GoTo ErrorTrap
  
  Dim fLocalRequired As Boolean
  Dim sSQL As String
  Dim rsData As ADODB.Recordset
  Dim lngExprID As Long
  
  fLocalRequired = False
  lngExprID = 0
  
  sSQL = "SELECT expressionID" & _
    " FROM ASRSysViews" & _
    " WHERE viewID = " & CStr(plngViewID)
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
    lngExprID = rsData!ExpressionID
  End If

  rsData.Close
  Set rsData = Nothing

  If lngExprID > 0 Then
    ' View does have a filter so now check if it uses any Hierarchy functions.
    fLocalRequired = HasFunctionComponent(lngExprID, 66)
    
    If Not fLocalRequired Then
      fLocalRequired = HasFunctionComponent(lngExprID, 68)
    End If
    
    If Not fLocalRequired Then
      fLocalRequired = HasFunctionComponent(lngExprID, 70)
    End If
    
    If Not fLocalRequired Then
      fLocalRequired = HasFunctionComponent(lngExprID, 72)
    End If
  End If
  
TidyUpAndExit:
  ViewRequiresLocalCursor = fLocalRequired
  Exit Function
  
ErrorTrap:
  fLocalRequired = False
  GoTo TidyUpAndExit
  
End Function
Public Function ChildViewRequiresLocalCursor(plngTableID As Long) As Boolean
  ' Return TRUE if the child view for the given table requires a LOCAL cursor.
  ' ie. if it one of its ascendent tables or view requires a local cursor.
  
  On Error GoTo ErrorTrap
  
  Dim fLocalRequired As Boolean
  Dim sSQL As String
  Dim rsData As ADODB.Recordset
  
  fLocalRequired = False
  
  sSQL = "SELECT ASRSysChildViewParents2.parentID," & _
    "   ASRSysChildViewParents2.parentType," & _
    "   ASRSysChildViewParents2.parentTableID" & _
    " FROM ASRSysChildViewParents2" & _
    " INNER JOIN ASRSysChildViews2 ON ASRSysChildViewParents2.childViewID = ASRSysChildViews2.childViewID" & _
    " WHERE ASRSysChildViews2.role = '" & gsUserGroup & "'" & _
    "   AND ASRSysChildViews2.tableID = " & CStr(plngTableID) & _
    "   AND (ASRSysChildViewParents2.parentType = 'UV'" & _
    "     OR ASRSysChildViewParents2.parentType = 'SV')" & _
    " ORDER BY ASRSysChildViewParents2.parentType"

  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Do Until rsData.EOF
    If Trim(rsData!ParentType) = "UV" Then
      fLocalRequired = ViewRequiresLocalCursor(rsData!ParentID)
    Else
      fLocalRequired = ChildViewRequiresLocalCursor(rsData!ParentTableID)
    End If
  
    If fLocalRequired Then Exit Do
    rsData.MoveNext
  Loop

  rsData.Close
  Set rsData = Nothing
  
TidyUpAndExit:
  ChildViewRequiresLocalCursor = fLocalRequired
  Exit Function
  
ErrorTrap:
  fLocalRequired = False
  GoTo TidyUpAndExit

End Function

Private Sub Class_Initialize()

  Set datData = New DataMgr.clsDataAccess

End Sub

Private Sub Class_Terminate()
  
  Set datData = Nothing

End Sub

Public Function GetPicture(plngPictureID As Long) As ADODB.Recordset

  Dim sSQL As String
  Dim rsPicture As ADODB.Recordset
    
  sSQL = "SELECT picture" & _
    " FROM ASRSysPictures" & _
    " WHERE pictureID = " & Trim(Str(plngPictureID))
    
  Set rsPicture = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetPicture = rsPicture

End Function

Public Function GetUserDetails() As String

  Dim sSQL As String
  Dim rsUser As ADODB.Recordset
  
  'JPD 20060106 Fault 10456
  ' JPD20020806 Fault 3901
  sSQL = "exec sp_helpuser '" & Replace(gsSQLUserName, "'", "''") & "'"
  Set rsUser = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  'MH20040114 Fault 5627
  'If rsUser!groupname = "db_owner" Then
  '    rsUser.MoveNext
  'End If
  Do While rsUser!groupname = "db_owner" _
        Or LCase(Left(rsUser!groupname, 6)) = "asrsys"
    rsUser.MoveNext
  Loop
  
  GetUserDetails = rsUser!groupname
  
  rsUser.Close
  Set rsUser = Nothing
  
End Function

Public Function GetSqlVersion() As String

  Dim sSQL As String
  Dim rsUser As ADODB.Recordset
  
  sSQL = "exec sp_server_info 2"
  
  Set rsUser = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  'rsUser.MoveNext
'  GetSqlVersion = "SQL Server Version: " & rsUser(3)
  GetSqlVersion = rsUser(2)

  rsUser.Close
  Set rsUser = Nothing
    
  If InStr(GetSqlVersion, vbLf) > 0 Then
    GetSqlVersion = Left(GetSqlVersion, InStr(GetSqlVersion, vbLf) - 1)
  End If
  
End Function

Public Sub DeleteRecord(psTableName As String, psIDField As String, plngID As Long)

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.DeleteRecord(psTableName,psIDField,plngID)", Array(psTableName, psIDField, plngID)
  
  ' Delete the given record.
  Dim sSQL As String
    
  sSQL = "DELETE FROM " & psTableName & " WHERE " & psIDField & " = " & plngID
  datData.ExecuteSql sSQL
       
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Public Function GetRecords(sSQL As String) As ADODB.Recordset
  ' Return the required forward-only/read-only recordset.
'  Set GetRecords = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  ' JPD - Changed cursor-type to static.
  ' The SQL Server ODBC driver uses a special cursor when the cursor is
  ' forward-only, read-only, and the ODBC rowset size is one.
  ' The cursor is called a "firehose" cursor because it is the fastest way to retrieve the data.
  ' Unfortunately, a side affect of the cursor is that it only permits one active recordset per connection.
  ' To get around this we'll try using a STATIC cursor.
  Set GetRecords = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)

End Function



Public Function GetPersistentRecords(sSQL As String, CursorType As ADODB.CursorTypeEnum, _
  LockType As ADODB.LockTypeEnum) As ADODB.Recordset
  Set GetPersistentRecords = datData.OpenPersistentRecordset(sSQL, CursorType, LockType)

End Function


Public Function GetRecordsInTransaction(sSQL As String) As ADODB.Recordset
  ' Return the required STATIC/read-only recordset.
  ' This is useful when getting a recordset in the middle of a transaction.
  ' An error occurs when getting more than one forward-only, read-only recordset in the middle
  ' of a transaction.
  Set GetRecordsInTransaction = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)

End Function



Public Function GetReadOnlyRecords(sSQL As String) As ADODB.Recordset

  'MH20030204 If the SQL causes an error then this will be handled
  'in "datData.OpenRecordset".  Having the error trapping in here
  'meant that I got the SQL string twice in the error message box.


  On Error GoTo ErrorTrap
  'gobjErrorStack.PushStack "clsGeneral.GetReadOnlyRecords(sSQL)", Array(sSQL)
  ' Return the required dynamic/read-only recordset.
'  Set GetReadOnlyRecords = datData.OpenRecordset(sSQL, adOpenDynamic, adLockReadOnly)

  ' JPD 7/6/00 Changed the cursor type from dynamic to static.
  ' A dynamic/read-only recordset opened with a sql select statement that includes the 'distinct' parameter,
  ' and returns a single record will automatically be set to be a forwardonly/read-only (firehose) recordet.
  Set GetReadOnlyRecords = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)

TidyUpAndExit:
  'gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  'gobjErrorStack.HandleError

End Function


Public Function GetPersistentMainRecordset(sSQL As String) As ADODB.Recordset
  ' Return the recordset for the given SQL select statement.
  ' The recordset uses an OPTIMISTICALLY locked KEYSET cursor.
  ' ie. the kind used for record editing.
  Set GetPersistentMainRecordset = datData.OpenPersistentRecordset(sSQL, adOpenKeyset, adLockOptimistic)
  
End Function

Public Function GetMainRecordset(sSQL As String) As ADODB.Recordset
  ' Return the recordset for the given SQL select statement.
  ' The recordset uses an OPTIMISTICALLY locked KEYSET cursor.
  ' ie. the kind used for record editing.
  Set GetMainRecordset = datData.OpenRecordset(sSQL, adOpenKeyset, adLockOptimistic)
  
End Function


Public Function GetOrders(plngTableID As Long) As ADODB.Recordset
  ' Return a recordset of the orders defined for the given table.
  Dim sSQL As String
  Dim rsOrders As ADODB.Recordset
    
  sSQL = "SELECT *" & _
    " FROM ASRSysOrders" & _
    " WHERE tableID = " & Trim(Str(plngTableID))
  Set rsOrders = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetOrders = rsOrders

  Set rsOrders = Nothing
  
End Function

Public Function GetRuntimeOrders(plngTableID As Long) As ADODB.Recordset
  ' Return a recordset of the RUNTIME orders defined for the given table.
  Dim sSQL As String
  Dim rsOrders As ADODB.Recordset
    
  sSQL = "SELECT *" & _
    " FROM ASRSysOrders" & _
    " WHERE tableID = " & Trim(Str(plngTableID)) & _
    " AND type = " & Trim(Str(giORDERTYPE_DYNAMIC))
  Set rsOrders = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetRuntimeOrders = rsOrders

  Set rsOrders = Nothing
  
End Function


Public Function GetRuntimeViewOrders(plngViewID As Long, plngTableID As Long) As ADODB.Recordset
  ' Return a recordset of the RUNTIME orders defined for the given view.
  Dim sSQL As String
  Dim rsOrders As ADODB.Recordset
    
  sSQL = "SELECT DISTINCT ASRSysOrders.OrderID, ASRSysOrders.Name, ASRSysOrders.TableID" & _
    " FROM ASRSysOrders" & _
    " INNER JOIN ASRSysOrderItems" & _
    "   ON ASRSysOrders.OrderID = ASRSysOrderItems.OrderID" & _
    " INNER JOIN ASRSysViewColumns" & _
    "   ON ASRSysOrderItems.ColumnID = ASRSysViewColumns.ColumnID" & _
    " WHERE ASRSysOrders.TableID = " & Trim(Str(plngTableID)) & _
    "   AND ASRSysViewColumns.ViewID = " & Trim(Str(plngViewID)) & _
    "   AND ASRSysViewColumns.InView = 1" & _
    "   AND ASRSysOrders.type = " & Trim(Str(giORDERTYPE_DYNAMIC))
  Set rsOrders = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetRuntimeViewOrders = rsOrders

  Set rsOrders = Nothing
  
End Function

Public Function GetViewOrders(plngViewID As Long, plngTableID As Long) As ADODB.Recordset
  
  Dim sSQL As String
  Dim rsOrders As ADODB.Recordset
    
  sSQL = "EXEC dbo.sp_ASRGetOrders " & plngViewID & ", " & plngTableID
  Set rsOrders = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetViewOrders = rsOrders

  Set rsOrders = Nothing
  
End Function

Public Function GetControlWidth(plngTableID As Long, plngColumnID As Long) As Long
  
  ' Return the width of the given column's control.
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
    
  sSQL = "SELECT width" & _
    " FROM ASRSysControls" & _
    " WHERE tableID = " & plngTableID & _
    " AND columnID = " & plngColumnID
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    GetControlWidth = rsTemp(0)
  Else
    GetControlWidth = 0
  End If
    
  rsTemp.Close
  Set rsTemp = Nothing
    
End Function

Public Function GetHistorySummaryFields(plngHistoryTableID As Long, plngParentTableID As Long) As clsSummaryFields

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetHistorySummaryFields(plngHistoryTableID,plngParentTableID)", Array(plngHistoryTableID, plngParentTableID)
  
  ' Return the recordset of summary fields for the given table.
  Dim iLoop As Integer
  Dim sSQL As String
  Dim rsSummary As ADODB.Recordset
  Dim objSummaryField As clsSummaryField
  Dim objSummaryFields As clsSummaryFields
    
  ' Instantiate the Summary Fields collection if it does not already exist.
  If gcolSummaryFieldsCollection Is Nothing Then
    Set gcolSummaryFieldsCollection = New Collection
  End If
    
  ' If the given Summary Fields collection has already been
  ' read then simply return it.
  For iLoop = 1 To gcolSummaryFieldsCollection.Count
    If gcolSummaryFieldsCollection(iLoop).Tag = Trim(Str(plngHistoryTableID)) & "_" & Trim(Str(plngParentTableID)) Then
      Set GetHistorySummaryFields = gcolSummaryFieldsCollection(iLoop)
      GoTo TidyUpAndExit
    End If
  Next iLoop
    
  Set objSummaryFields = New clsSummaryFields
  objSummaryFields.Tag = Trim(Str(plngHistoryTableID)) & "_" & Trim(Str(plngParentTableID))
    
  sSQL = "EXEC dbo.sp_ASRGetSummaryFields " & Trim(Str(plngHistoryTableID)) & ", " & Trim(Str(plngParentTableID))
  Set rsSummary = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Do While Not rsSummary.EOF
    Set objSummaryField = objSummaryFields.Add(rsSummary!Sequence)
    
    With objSummaryField
      .Sequence = rsSummary!Sequence
      
      ' JDM - Fault 2463 - Display breaks
      '.StartOfGroup = (rsSummary!StartOfGroup = 1)
      .StartOfGroup = rsSummary!StartOfGroup
      
      .ColumnName = rsSummary!ColumnName
      .ColumnID = rsSummary!ColumnID
      .TableID = rsSummary!TableID
      .DataType = rsSummary!DataType
      .Size = rsSummary!Size
      .Decimals = rsSummary!Decimals
      .ControlType = rsSummary!ControlType
      .ColumnType = rsSummary!ColumnType
      .Multiline = (rsSummary!Multiline = 1)
      .Alignment = rsSummary!Alignment
      .Use1000Separator = IIf(IsNull(rsSummary!Use1000Separator), False, rsSummary!Use1000Separator)
      .BlankIfZero = IIf(IsNull(rsSummary!BlankIfZero), False, rsSummary!BlankIfZero)
      
      'JDM - 22/11/01 - Fault 3175 - Added manual column break functionality
      .StartOfColumn = IIf(IsNull(rsSummary!StartOfColumn), False, rsSummary!StartOfColumn)
      
    End With
    
    rsSummary.MoveNext
  Loop
  rsSummary.Close
  Set rsSummary = Nothing
    
  ' Set if these summary fields have a manual / automatic column break
  objSummaryFields.ManualColumnBreak = datGeneral.GetManualSummaryBreak(plngHistoryTableID)
    
  gcolSummaryFieldsCollection.Add objSummaryFields, Trim(Str(plngHistoryTableID)) & "_" & Trim(Str(plngParentTableID))
    
  Set GetHistorySummaryFields = objSummaryFields

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

Public Function GetTableName(ByVal plngTableID As Long) As String
  Dim rsTable As ADODB.Recordset
  Dim sSQL As String
    
  sSQL = "SELECT tableName " & _
    " FROM ASRSysTables " & _
    " WHERE tableID=" & plngTableID
        
  Set rsTable = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsTable
    If Not (.BOF And .EOF) Then
      GetTableName = Trim(.Fields(0).Value)
    Else
      GetTableName = vbNullString
    End If
    
    .Close
  End With
    
  Set rsTable = Nothing
  
End Function

Public Function GetPicklistName(ByVal lPicklistID As Long) As String

  Dim rsPicklist As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = "SELECT Name " & _
      "FROM ASRSysPicklistName " & _
      "WHERE PicklistID=" & lPicklistID
      
  Set rsPicklist = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsPicklist
      If Not (.BOF And .EOF) Then
          GetPicklistName = Trim(.Fields(0).Value)
      Else
          GetPicklistName = vbNullString
      End If
      .Close
  End With
  
  Set rsPicklist = Nothing
  
End Function

Public Function GetFilterName(ByVal lFilterID As Long) As String
  
  Dim rsFilter As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = "SELECT name " & _
    "FROM ASRSysExpressions " & _
    "WHERE ExprID=" & lFilterID
      
  Set rsFilter = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsFilter
    If Not (.BOF And .EOF) Then
      GetFilterName = Trim(.Fields(0).Value)
    Else
      GetFilterName = vbNullString
    End If
    .Close
  End With
  
  Set rsFilter = Nothing
  
End Function

Public Function GetScreenName(ByVal ScreenID As Long) As String

  Dim rsScreen As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = "SELECT Name " & _
      "FROM ASRSysScreens " & _
      "WHERE ScreenID = " & ScreenID
      
  Set rsScreen = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsScreen
      If Not (.BOF And .EOF) Then
          GetScreenName = Trim(.Fields(0).Value)
      Else
          GetScreenName = vbNullString
      End If
      .Close
  End With
  
  Set rsScreen = Nothing
  
End Function

Public Function GetRecDescExprID(ByVal TableID As Long) As String
  
  ' JPD - Return the Record Description Expression ID for the given table.
  Dim rsTable As ADODB.Recordset
  Dim sSQL As String
    
  sSQL = "SELECT recordDescExprID " & _
    "FROM ASRSysTables " & _
    "WHERE TableID=" & TableID
        
  Set rsTable = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsTable
    If Not (.BOF And .EOF) Then
      GetRecDescExprID = .Fields(0).Value
    Else
      GetRecDescExprID = 0
    End If
    .Close
  End With
  
  Set rsTable = Nothing
  
End Function

Public Function GetADOConnection() As ADODB.Connection
  ' JPD - Return the ADO Connection object.
  Set GetADOConnection = gADOCon
  
End Function

Public Function GetViewTableName(ByVal lViewID As Long) As String

  Dim rsView As ADODB.Recordset
  Dim sSQL As String
  
  sSQL = "SELECT ViewName " & _
      "FROM ASRSysViews " & _
      "WHERE ViewID=" & lViewID
      
  Set rsView = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  With rsView
    If Not (.BOF And .EOF) Then
      GetViewTableName = Trim(.Fields(0).Value)
    Else
      GetViewTableName = vbNullString
    End If
  End With
  
  rsView.Close
  Set rsView = Nothing
  
End Function

Public Function GetTableIDFromTableViewName(ByVal psTableViewName As String) As Long
  ' Return the table ID of the given table/view.
  Dim rsTableInfo As ADODB.Recordset
  Dim sSQL As String
    
  sSQL = "SELECT tableID" & _
    " FROM ASRSysTables" & _
    " WHERE tableName = '" & psTableViewName & "'" & _
    " UNION " & _
    " SELECT viewTableID" & _
    " FROM ASRSysViews" & _
    " WHERE viewName = '" & psTableViewName & "'"
  Set rsTableInfo = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsTableInfo
    If Not (.BOF And .EOF) Then
      GetTableIDFromTableViewName = .Fields(0).Value
    Else
      GetTableIDFromTableViewName = 0
    End If
  End With
    
  rsTableInfo.Close
  Set rsTableInfo = Nothing
  
End Function

Public Function RecordAmended(plngTableID As Long, _
  psRealSource As String, _
  plngRecordID As Long, _
  plngTimestamp As Long) As Integer

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.RecordAmended2(psRealSource,plngRecordID,plngTimestamp)", Array(psRealSource, plngRecordID, plngTimestamp)
  
  ' Return 0 if the given record has NOT been amended.
  ' Return 1 if the given record has been amended AND is still in the given table/view.
  ' Return 2 if the given record has been amended AND is no longer in the given table/view.
  ' Return 3 if the given record has been deleted from the table.
  Dim fDeleted As Boolean
  Dim iResult As Integer
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fDoneOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sErrorMsg As String
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC
    
  Const iRETRIES = 5
  Const iPAUSE = 5000
    
  iOldCursorLocation = gADOCon.CursorLocation
  fDoneOK = True
  iRetryCount = 0
  
  Const sSPName = "sp_ASRRecordAmended"
  
  iResult = 0
  
  If plngRecordID > 0 Then
        
      ' Run the stored procedure to see if the given record has changed
      Set cmADO = New ADODB.Command
      With cmADO
        .CommandText = sSPName
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = gADOCon
              
        Set pmADO = .CreateParameter("resultCode", adInteger, adParamOutput)
        .Parameters.Append pmADO
            
        Set pmADO = .CreateParameter("tableID", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = plngTableID
        
        Set pmADO = .CreateParameter("realSource", adLongVarChar, adParamInput, -1)
        .Parameters.Append pmADO
        pmADO.Value = psRealSource
              
        Set pmADO = .CreateParameter("id", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = plngRecordID
              
        Set pmADO = .CreateParameter("timestamp", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = plngTimestamp
              
        Set pmADO = Nothing
        
        fDeadlock = True
        Do While fDeadlock
          fDeadlock = False
          
          ' Change the cursor location to 'client' as the errors that might be raised
          ' during the update cannot be read for 'server' cursors.
          gADOCon.Errors.Clear
          gADOCon.CursorLocation = adUseClient
                
          On Error GoTo DeadlockErrorTrap
DeadlockRecoveryPoint:
          cmADO.Execute
      
          On Error GoTo ErrorTrap

          ' Restore the original cursor location to the ADO connection object.
          gADOCon.CursorLocation = iOldCursorLocation
                
          ' Check if the update prodcued any errors.
          If gADOCon.Errors.Count > 0 Then
            sErrorMsg = ""
          
            For Each ADOErr In gADOCon.Errors
              ' If any 'deadlocks' occur, try to save changes again.
              ' Do this a few times and if errors still occur then display a more friendly
              ' error message than the ' deadlock victim' one generated by ODBC.
              If (ADOErr.Number = DEADLOCK_ERRORNUMBER) And _
                (((UCase(Left(ADOErr.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                  (UCase(Right(ADOErr.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                ((UCase(Left(ADOErr.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                  (InStr(UCase(ADOErr.Description), DEADLOCK2_MESSAGEEND) > 0))) Then

                ' The error is for a deadlock.
                ' Sorry about having to use the err.description to trap the error but the err.number
                ' is not specific and MSDN suggests using the err.description.
                If (iRetryCount < iRETRIES) And (gADOCon.Errors.Count = 1) Then
                  iRetryCount = iRetryCount + 1
                  fDeadlock = True
                  ' Pause before resubmitting the SQL command.
                  Sleep iPAUSE
                Else
                  sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                    "Another user is deadlocking the database. Try saving again."
                  fDoneOK = False
                End If
              Else
                sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                  ADOErr.Description
                fDoneOK = False
              End If
            Next ADOErr
          
            gADOCon.Errors.Clear
        
            If Not fDoneOK Then
              COAMsgBox "ERROR." & vbCrLf & vbCrLf & _
                sErrorMsg, vbOKOnly + vbExclamation, app.ProductName
            End If
          End If

          If fDoneOK And (Not fDeadlock) Then
            iResult = .Parameters(0).Value
          End If
        Loop
      End With
      Set cmADO = Nothing

  End If
  
  RecordAmended = iResult
  
TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    gADOCon.CursorLocation = iOldCursorLocation
  Else
    gADOCon.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  
  gobjErrorStack.PopStack
  Exit Function

ErrorTrap:
  fDoneOK = False
  gobjErrorStack.HandleError
  Exit Function
  
DeadlockErrorTrap:
  ' If any 'deadlocks' occur, try to save changes again.
  ' Do this a few times and if errors still occur then display a more friendly
  ' error message than the 'deadlock victim' one generated by ODBC.
  If (Err.Number = DEADLOCK_ERRORNUMBER) And _
    (((UCase(Left(Err.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
      (UCase(Right(Err.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
    ((UCase(Left(Err.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
      (InStr(UCase(Err.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    ' The error is for a deadlock.
    ' Sorry about having to use the err.description to trap the error but the err.number
    ' is not specific and MSDN suggests using the err.description.
    If iRetryCount < iRETRIES Then
      iRetryCount = iRetryCount + 1
      ' Pause before resubmitting the SQL command.
      Sleep iPAUSE
      Resume DeadlockRecoveryPoint
    Else
      fDoneOK = False
      COAMsgBox "Another user is deadlocking the database. Try saving again.", _
        vbExclamation + vbOKOnly, Application.Name
      gobjErrorStack.HandleError
      Resume TidyUpAndExit
    End If
  Else
    fDoneOK = False
    COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
    gobjErrorStack.HandleError
    Resume TidyUpAndExit
  End If

End Function

Public Function GetDataType(lTableID As Long, lColumnID As Long) As Long

  Dim sSQL As String
  Dim rsData As ADODB.Recordset

  sSQL = "Select datatype From ASRSysColumns Where columnid= " & lColumnID & " And tableID = " & lTableID
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
    GetDataType = rsData(0).Value
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Public Function GetColumnDataType(lColumnID As Long) As Long

  Dim sSQL As String
  Dim rsData As ADODB.Recordset

  sSQL = "Select datatype From ASRSysColumns Where columnid = " & lColumnID
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
    GetColumnDataType = rsData(0)
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Public Function GetColumnTable(plngColumnID As Long) As Long
  
  ' Return the table id of the given column.
  Dim sSQL As String
  Dim rsData As ADODB.Recordset

  sSQL = "SELECT tableID" & _
    " FROM ASRSysColumns" & _
    " WHERE columnID = " & Trim(Str(plngColumnID))
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
    GetColumnTable = rsData!TableID
  Else
    GetColumnTable = 0
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Public Function GetDataSize(lColumnID As Long, Optional blnNotZero As Boolean) As Long

    Dim sSQL As String
    Dim rsTemp As ADODB.Recordset

    'sSQL = "Select size From ASRSysColumns Where columnid = " & lColumnID & " And tableID = " & lTableID
    sSQL = "Select size, datatype From ASRSysColumns Where columnid = " & lColumnID
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

    'If Not rsTemp.BOF And Not rsTemp.EOF Then
    '    GetDataSize = rsTemp(0)
    'End If
    If Not rsTemp.BOF And Not rsTemp.EOF Then
      If rsTemp!DataType = sqlDate Or rsTemp!DataType = sqlInteger Then
        GetDataSize = 10
      ElseIf rsTemp!DataType = sqlLongVarChar Then  'Working Pattern
        GetDataSize = 14
      Else
        GetDataSize = rsTemp(0)
      End If
    End If
    
    
    If blnNotZero Then
      If rsTemp!DataType = sqlBoolean Then
        GetDataSize = 5   'Allow "False"
      End If
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

End Function

'Public Function RunDataTransfer(plTransferID As Long, bSilentmode As Boolean) As Boolean
' NO LONGER USED.
'  ' Perform the given Data Transfer.
'  On Error GoTo Err_Trap
'
'  Dim bFinished As Boolean
'  Dim bNoPrimary As Boolean
'  Dim lPos As Long
'  Dim lCount As Long
'  Dim lAffected As Long
'  Dim lRecordID As Long
'  Dim lRecordCount As Long
'  Dim lCurrentTableID As Long
'  Dim sSQL As String
'  Dim sFrom As String
'  Dim sTemp As String
'  Dim sSelect As String
'  Dim sToTableName As String
'  Dim sFilteredIDs As String
'  Dim sCurrentTable As String
'  Dim sFromTableName As String
'  'Dim WaitWindow As NewWaitMsg.clsNewWaitMsg
'  Dim rsTemp As ADODB.Recordset
'  Dim rsWhere As ADODB.Recordset
'  Dim rsJoin As ADODB.Recordset
'  Dim adoConnect As ADODB.Connection
'  Dim sInsert() As String
'
'  Screen.MousePointer = vbHourglass
'  'Set WaitWindow = New NewWaitMsg.clsNewWaitMsg
'  'WaitWindow.Initialise 8, "Transferring data...", App.Title, True, False, True, App.Path & "\videos\transfer.avi"
'
'  ' Get the details of the given Data Transfer.
'  sSQL = "EXEC sp_ASRWhereTransferDetails " & plTransferID
'  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'
'  'WaitWindow.UpdateProgress
'
'  ' Get a recordset of IDs of the records to be transferred.
'  If rsTemp!AllRecords Then
'    ' Get ALL the record id's from the given table.
'    sSQL = "SELECT id FROM " & rsTemp!TableName
'    Set rsWhere = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
'    rsTemp.Close
'
'  ElseIf rsTemp!FilterID > 0 Then
'    ' Get a FILTER selection of record id's from the given table.
'    sSQL = "SELECT id FROM " & rsTemp!TableName
'    sFilteredIDs = GetFilteredIDs(rsTemp!FilterID)
'    If Len(sFilteredIDs) > 0 Then
'      sSQL = sSQL & " WHERE id IN (" & sFilteredIDs & ")"
'    End If
'    Set rsWhere = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
'    rsTemp.Close
'
'  Else
'    ' Get a PICKLIST selection of record id's from the given table.
'    sSQL = "EXEC sp_ASRGetPickListRecords " & rsTemp!PickListID
'    Set rsWhere = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
'    rsTemp.Close
'
'  End If
'
'  'WaitWindow.UpdateProgress
'
'  ' Check if we have any records to transfer.
'  If rsWhere.BOF And rsWhere.EOF Then
'    'Set WaitWindow = Nothing
'
'    'COAMsgBox "No records to transfer from selected option.", vbExclamation, "Data Transfer"
'    If bSilentmode = False Then COAMsgBox "No records to transfer from selected option.", vbExclamation, "Data Transfer"
'
'    rsWhere.Close
'    Set rsWhere = Nothing
'    Set rsTemp = Nothing
'    Screen.MousePointer = vbDefault
'    RunDataTransfer = False
'    Exit Function
'  End If
'
'  ' Get the primary table insert parameters.
'  sSQL = "EXEC sp_ASRPrimaryTransferDetails " & plTransferID
'  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'
'  'WaitWindow.UpdateProgress
'
'  ' Construct the primary INSERT SQL string.
'  ReDim Preserve sInsert(1)
'  If Not rsTemp.BOF And Not rsTemp.EOF Then
'
'    sToTableName = rsTemp!ToTableName
'    sFromTableName = rsTemp!FromPrimaryTable
'
'    sInsert(1) = "INSERT " & rsTemp!ToTableName & " ("
'    sSelect = "SELECT "
'    sFrom = " FROM " & rsTemp!FromPrimaryTable
'
'    Do While Not rsTemp.EOF
'      sInsert(1) = sInsert(1) & rsTemp!ToColumnName
'      If rsTemp!FromSysDate Then
'        sSelect = sSelect & "'" & Format(Now, "mm/dd/yy") & "'"
'      ElseIf IsNull(rsTemp!FromTableName) Then
'        sSelect = sSelect & "'" & rsTemp!FromText & "'"
'      Else
'        sSelect = sSelect & rsTemp!FromTableName & "." & rsTemp!FromColumnName
'      End If
'      rsTemp.MoveNext
'      If Not rsTemp.EOF Then
'        sInsert(1) = sInsert(1) & ", "
'        sSelect = sSelect & ", "
'      Else
'        sInsert(1) = sInsert(1) & ")"
'      End If
'    Loop
'    rsTemp.Close
'    sInsert(1) = sInsert(1) & " " & sSelect & sFrom
'
'    sSQL = "EXEC sp_ASRPrimaryJoinDetails " & plTransferID
'    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'    If Not rsTemp.BOF And Not rsTemp.EOF Then
'      sInsert(1) = sInsert(1) & " INNER JOIN "
'    End If
'    Do While Not rsTemp.EOF
'      sInsert(1) = sInsert(1) & rsTemp!FromTableName & " ON " & rsTemp!FromPrimaryTable & _
'        ".id = " & rsTemp!FromTableName & ".id_" & rsTemp!FromPrimaryTableID
'      rsTemp.MoveNext
'      If Not rsTemp.EOF Then
'        sInsert(1) = sInsert(1) & " INNER JOIN "
'      End If
'    Loop
'    rsTemp.Close
'  Else
'    rsTemp.Close
'    bNoPrimary = True
'    sSQL = "SELECT ASRSysTables.TableName AS PrimaryTableName" & _
'      " FROM ASRSysDataTransferName" & _
'      " INNER JOIN ASRSysTables ON ASRSysDataTransferName.FromTableID = ASRSysTables.TableID" & _
'      " WHERE ASRSysDataTransferName.DataTransferID = " & plTransferID
'    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'    sFromTableName = rsTemp!PrimaryTableName
'    rsTemp.Close
'  End If
'
'  'WaitWindow.UpdateProgress
'  ' Get the secondary table insert parameters.
'  sSQL = "EXEC sp_ASRSecondaryTransferDetails " & plTransferID
'  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  lCount = 1
'  bFinished = True
'
'  'WaitWindow.UpdateProgress
'
'  ' Construct the secondary INSERT SQL string.
'  Do While Not rsTemp.EOF
'    If bFinished Then
'      lCount = lCount + 1
'      sSelect = "SELECT "
'      ReDim Preserve sInsert(lCount)
'      sInsert(lCount) = "INSERT " & rsTemp!ToTableName & " ("
'      sFrom = " FROM " & rsTemp!FromPrimaryTable
'    End If
'
'    sInsert(lCount) = sInsert(lCount) & rsTemp!ToColumnName
'    If rsTemp!FromSysDate Then
'      sSelect = sSelect & "'" & Format(Now, "mm/dd/yy") & "'"
'    ElseIf IsNull(rsTemp!FromTableName) Then
'      sSelect = sSelect & "'" & rsTemp!FromText & "'"
'    Else
'      sSelect = sSelect & rsTemp!FromTableName & "." & rsTemp!FromColumnName
'    End If
'
'    sCurrentTable = rsTemp!ToTableName
'    lCurrentTableID = rsTemp!ToPrimaryTableID
'
'    rsTemp.MoveNext
'    If Not rsTemp.EOF Then
'      If rsTemp!ToTableName = sCurrentTable Then
'        bFinished = False
'        sInsert(lCount) = sInsert(lCount) & ", "
'        sSelect = sSelect & ", "
'      Else
'        bFinished = True
'        sInsert(lCount) = sInsert(lCount) & ", id_" & lCurrentTableID
'        sSelect = sSelect & ", *****"
'        sInsert(lCount) = sInsert(lCount) & ") " & sSelect & sFrom
'
'        sSQL = "EXEC sp_ASRSecondaryJoinDetails " & plTransferID & ", '" & sCurrentTable & "'"
'        Set rsJoin = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'        If Not rsJoin.BOF And Not rsJoin.EOF Then
'          sInsert(lCount) = sInsert(lCount) & " INNER JOIN "
'        End If
'        Do While Not rsJoin.EOF
'          sInsert(lCount) = sInsert(lCount) & rsJoin!FromTableName & " ON " & rsJoin!FromPrimaryTable & _
'                ".id = " & rsJoin!FromTableName & ".id_" & rsJoin!FromPrimaryTableID
'          rsJoin.MoveNext
'          If Not rsJoin.EOF Then
'            sInsert(lCount) = sInsert(lCount) & " INNER JOIN "
'          End If
'        Loop
'        rsJoin.Close
'      End If
'    Else
'      sInsert(lCount) = sInsert(lCount) & ", id_" & lCurrentTableID
'      sSelect = sSelect & ", *****"
'      sInsert(lCount) = sInsert(lCount) & ") " & sSelect & sFrom
'
'      sSQL = "EXEC sp_ASRSecondaryJoinDetails " & plTransferID & ", '" & sCurrentTable & "'"
'      Set rsJoin = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'      If Not rsJoin.BOF And Not rsJoin.EOF Then
'        sInsert(lCount) = sInsert(lCount) & " INNER JOIN "
'      End If
'      Do While Not rsJoin.EOF
'        sInsert(lCount) = sInsert(lCount) & rsJoin!FromTableName & " ON " & rsJoin!FromPrimaryTable & _
'          ".id = " & rsJoin!FromTableName & ".id_" & rsJoin!FromPrimaryTableID
'        rsJoin.MoveNext
'        If Not rsJoin.EOF Then
'          sInsert(lCount) = sInsert(lCount) & " INNER JOIN "
'        End If
'      Loop
'      rsJoin.Close
'    End If
'  Loop
'  rsTemp.Close
'
'  'Now actually Insert the records, primary first(getting the recordid back), then secondary
'  'First set up the progress form progress bar
'  rsWhere.MoveLast
'  rsWhere.MoveFirst
'
'  'WaitWindow.UpdateProgress
'
'  'Make a fresh connection to SQL for this operation to speed things up a bit...
'  Set adoConnect = New ADODB.Connection
'  With adoConnect
'    .ConnectionString = gsConnectString
'    .Provider = "SQLOLEDB"
'    .CommandTimeout = 0
'    .CursorLocation = adUseServer
'    .Mode = adModeReadWrite
'    .Open
'  End With
'
'  'WaitWindow.UpdateProgress
'
'  Do While Not rsWhere.EOF
'    'Do the primary insert to get back the recently added recordid, if there is a primary...
'    If Not bNoPrimary Then
'      sTemp = sInsert(1)
'      sTemp = sTemp & " WHERE " & sFromTableName & ".id = " & rsWhere!ID
'      adoConnect.Execute sTemp, lAffected, adCmdText
'      sSQL = "SELECT MAX(id) FROM " & sToTableName
'      rsTemp.Open sSQL, adoConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
'      lRecordID = rsTemp(0)
'      rsTemp.Close
'    Else
'      lRecordID = rsWhere(0)
'    End If
'
'    If UBound(sInsert, 1) > 1 Then
'      'Then do all the secondaries
'      For lCount = 2 To UBound(sInsert, 1)
'        sTemp = sInsert(lCount)
'        sTemp = sTemp & " WHERE " & sFromTableName & ".id = " & rsWhere!ID
'        lPos = InStr(1, sTemp, "*****")
'        sTemp = Mid$(sTemp, 1, lPos - 1) & lRecordID & _
'          Mid$(sTemp, lPos + 5, Len(sTemp))
'        adoConnect.Execute sTemp, lAffected, adCmdText
'      Next
'    End If
'
'    rsWhere.MoveNext
'    lRecordCount = lRecordCount + lAffected
'  Loop
'
''WaitWindow.UpdateProgress
'
'  rsWhere.Close
'  Set rsWhere = Nothing
'  Set rsTemp = Nothing
'  Set rsJoin = Nothing
'
'  adoConnect.Close
'  Set adoConnect = Nothing
'
'  'Set WaitWindow = Nothing
'
'  Screen.MousePointer = vbDefault
'  If lRecordCount > 1 Then
'    sTemp = lRecordCount & " Records Transfered."
'  ElseIf lRecordCount = 0 Then
'    sTemp = "No Records Transfered."
'  Else
'    sTemp = lRecordCount & " Record Transfered."
'  End If
'
'  If bSilentmode = False Then COAMsgBox sTemp, vbInformation, "Data Transfer"
'
'  RunDataTransfer = True
'
'  Exit Function
'
'Err_Trap:
''Set WaitWindow = Nothing
'
'  If bSilentmode = False Then COAMsgBox Err.Description
'
'  Screen.MousePointer = vbDefault
'  RunDataTransfer = False
'
'End Function

Public Function GetDefaultOrder(plngTableID As Long) As Long

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetDefaultOrder(plngTableID)", Array(plngTableID)
  
  ' Return the default order ID for the given table.
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
    
  sSQL = "SELECT defaultOrderID" & _
    " FROM ASRSysTables" & _
    " WHERE tableID = " & Trim(Str(plngTableID))
  Set rsInfo = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If Not (rsInfo.BOF And rsInfo.EOF) Then
    GetDefaultOrder = rsInfo(0)
  Else
    GetDefaultOrder = 0
  End If
  rsInfo.Close
  Set rsInfo = Nothing

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

Public Function GetOrder(lOrderID As Long) As ADODB.Recordset

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetOrder(lOrderID)", Array(lOrderID)
    
  Dim sSQL As String
  Dim rsOrder As ADODB.Recordset
      
  sSQL = "SELECT * " & _
      "FROM ASRSysOrders " & _
      "WHERE OrderID=" & lOrderID
  Set rsOrder = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  Set GetOrder = rsOrder

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function



Public Function GetTableScreens() As ADODB.Recordset

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetTableScreens()"

  Dim sSQL As String
  Dim rsTableScreens As ADODB.Recordset
    
  sSQL = "SELECT ASRSysTables.tableID, ASRSysTables.tableName, ASRSysScreens.pictureID" & _
    " FROM ASRSysTables, ASRSysScreens" & _
    " WHERE ASRSysTables.tableType = " & Trim(Str(tabLookup)) & _
    " AND (ASRSysTables.tableID = ASRSysScreens.tableID)" & _
    " AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))" & _
    " AND ((ASRSysScreens.quickEntry IS null) OR (ASRSysScreens.quickEntry = 0))" & _
    " ORDER BY ASRSysTables.tableName DESC"
    
  Set rsTableScreens = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
  Set GetTableScreens = rsTableScreens

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

'Public Function GetSupportInfo() As ADODB.Recordset
'
'    Dim sSQL As String
'    Dim rsSupport As ADODB.Recordset
'
'    sSQL = "Select * From ASRSysConfig"
'    Set rsSupport = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'    Set GetSupportInfo = rsSupport
'
'End Function

Public Function ExecuteReturnAffected(sSQL As String) As Long

    ExecuteReturnAffected = datData.ExecuteSqlReturnAffected(sSQL)

End Function

'MH20020122 Not Used ?
'Public Function IsInView(plngViewID As Long, plngColumnID As Long) As Boolean
'
'  On Error GoTo ErrorTrap
'  gobjErrorStack.PushStack "clsGeneral.IsInView(plngViewID,plngColumnID)", Array(plngViewID, plngColumnID)
'
'  ' Return TRUE if the given column is in the given view.
'  Dim sSQL As String
'  Dim rsView As ADODB.Recordset
'
'  If plngViewID > 0 Then
'    sSQL = "SELECT inView" & _
'      " FROM ASRSysViewColumns" & _
'      " WHERE viewID = " & Trim(Str(plngViewID)) & _
'      " AND columnID = " & Trim(Str(plngColumnID))
'    Set rsView = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'
'    IsInView = rsView(0)
'    rsView.Close
'    Set rsView = Nothing
'  Else
'    IsInView = True
'  End If
'
'TidyUpAndExit:
'  gobjErrorStack.PopStack
'  Exit Function
'ErrorTrap:
'  gobjErrorStack.HandleError
'
'End Function
'
'MH20020122 Not Used ?
'Public Function GetLastID(psTableName As String) As Long
'
'  On Error GoTo ErrorTrap
'  gobjErrorStack.PushStack "clsGeneral.GetLastID(psTableName)", Array(psTableName)
'
'  ' Return the highest ID in the given table.
'  ' If no records match the given tablename then return -1.
'
'  Dim fOK As Boolean
'  Dim lngMaxID As Long
'  Dim sSQL As String
'  Dim rsTemp As ADODB.Recordset
'
'  fOK = True
'
'  sSQL = "SELECT MAX(id) FROM " & psTableName
'  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  fOK = Not IsNull(rsTemp(0))
'
'  If fOK Then
'    lngMaxID = rsTemp(0)
'  End If
'
'  rsTemp.Close
'
'TidyUpAndExit:
'  Set rsTemp = Nothing
'  If fOK Then
'    GetLastID = lngMaxID
'  Else
'    GetLastID = -1
'  End If
'
'  gobjErrorStack.PopStack
'  Exit Function
'ErrorTrap:
'  fOK = False
'  gobjErrorStack.HandleError
'
'End Function

Public Function GetOrderValue(plngID As Long, _
  psColumnName As String, _
  psTableName As String) As Variant
  ' Return the value from the given column for the given record.
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
    
  sSQL = "SELECT " & psColumnName & " FROM " & psTableName & " WHERE id = " & plngID
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If Not (rsTemp.EOF And rsTemp.BOF) Then
    GetOrderValue = rsTemp(0)
  Else
  
    ' JDM - 27/08/02 - Fault 4319 - Format the return type
    sSQL = "SELECT TOP 1 " & psColumnName & " FROM " & psTableName
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

    If Not (rsTemp.EOF And rsTemp.BOF) Then
      GetOrderValue = rsTemp(0)
    End If

  End If
  
  rsTemp.Close
  Set rsTemp = Nothing

End Function

'MH20020122 Not Used ?
'Public Function GetColumnDefault(sColumnName As String, lTableID As Long) As ADODB.Recordset
'
'  On Error GoTo ErrorTrap
'  gobjErrorStack.PushStack "clsGeneral.GetColumnDefault(sColumnName,lTableID)", Array(sColumnName, lTableID)
'
'  Dim sSQL As String
'  Dim rsTemp As ADODB.Recordset
'
'  sSQL = "EXEC sp_ASRColumnDefault '" & sColumnName & "'," & lTableID
'  Set rsTemp = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
'  Set GetColumnDefault = rsTemp
'
'TidyUpAndExit:
'  gobjErrorStack.PopStack
'  Exit Function
'ErrorTrap:
'  gobjErrorStack.HandleError
'
'End Function
'
'MH20020122 Not Used ?
'Public Function GetValidation(psColumnName As String, plngTableID As Long) As ADODB.Recordset
'
'  On Error GoTo ErrorTrap
'  gobjErrorStack.PushStack "clsGeneral.GetValidation(psColumnName,plngTableID)", Array(psColumnName, plngTableID)
'
'  ' Get the unique and mandatory constraints on the given column.
'  Dim sSQL As String
'  Dim rsValidate As ADODB.Recordset
'
'  sSQL = "exec sp_ASRUserValidation '" & psColumnName & "', " & plngTableID
'  Set rsValidate = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  Set GetValidation = rsValidate
'
'TidyUpAndExit:
'  gobjErrorStack.PopStack
'  Exit Function
'ErrorTrap:
'  gobjErrorStack.HandleError
'
'End Function

'MH20020122 Not Used ?
'Public Function UniqueCheck(psColumnName As String, plngTableID As Long) As ADODB.Recordset
'
'  On Error GoTo ErrorTrap
'  gobjErrorStack.PushStack "clsGeneral.UniqueCheck(psColumnName,plngTableID)", Array(psColumnName, plngTableID)
'
'  Dim sSQL As String
'  Dim rsUnique As ADODB.Recordset
'
'  sSQL = "exec sp_ASRUniqueCheck '" & psColumnName & "', " & plngTableID
'  Set rsUnique = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  Set UniqueCheck = rsUnique
'
'TidyUpAndExit:
'  gobjErrorStack.PopStack
'  Exit Function
'ErrorTrap:
'  gobjErrorStack.HandleError
'
'End Function

Public Function GetStatusBarText(sColName As String, lTableID As Long) As String

  Dim sSQL As String
  Dim rstText As ADODB.Recordset
  
  sSQL = "SELECT statusBarMessage FROM ASRSysColumns WHERE TableID = " & lTableID & " AND columnname = '" & sColName & "'"
  Set rstText = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If Not rstText.BOF And Not rstText.EOF Then
      GetStatusBarText = IIf(IsNull(rstText(0)), "", rstText(0))
  End If
  
  rstText.Close
  Set rstText = Nothing

End Function

Public Function GetQuickEntryScreens() As ADODB.Recordset
  Dim sSQL As String
  Dim rsScreens As ADODB.Recordset
    
  ' JPD20030115 Fault 4940
  sSQL = "SELECT ASRSysScreens.screenID, ASRSysScreens.name, ASRSysScreens.pictureID," & _
    " ASRSysTables.tableName, ASRSysTables.tableID" & _
    " FROM ASRSysScreens" & _
    " INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysScreens.quickEntry = 1" & _
    " ORDER BY ASRSysScreens.name DESC"
        
  Set rsScreens = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetQuickEntryScreens = rsScreens

End Function

Public Function GetColumnName(plngColumnID As Long, Optional pbIncludeTableName As Boolean = False) As String
  ' Return the name of the given column.
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset

  If pbIncludeTableName Then
    sSQL = "SELECT t.[tablename] + '.' + c.[columnName] FROM dbo.[ASRSysColumns] c " _
        & "INNER JOIN dbo.[ASRSysTables] t ON t.[TableID] = c.[TableID]  " _
        & "WHERE columnID = " & plngColumnID & ";"
  Else
    sSQL = "SELECT [columnName] FROM dbo.[ASRSysColumns] WHERE [columnID] = " & plngColumnID & ";"
  End If
  
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If (rsTemp.EOF And rsTemp.BOF) Then
    GetColumnName = vbNullString
  Else
    GetColumnName = rsTemp(0).Value
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing
    
End Function

Public Function GetColumnDefinition(plngColumnID As Long) As ADODB.Recordset

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetColumnDefinition(plngColumnID)", Array(plngColumnID)

  Dim sSQL As String
  Dim rsColumnDef As ADODB.Recordset

  sSQL = "SELECT * FROM ASRSysColumns WHERE columnID = " & Trim(Str(plngColumnID))
  Set rsColumnDef = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetColumnDefinition = rsColumnDef
    
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

Public Function GetExpression(lExprID As Long) As String

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetExpression(lExprID)", Array(lExprID)

  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
  
  sSQL = "Select Name From ASRSysExpressions Where ExprID = " & lExprID
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  If rsTemp.BOF And rsTemp.EOF Then
    GetExpression = vbNullString
  Else
    GetExpression = rsTemp(0)
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing
    
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function


Public Function GetEmailGroupName(lngEmailID As Long) As String

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetEmailGroupName(lngEmailID)", Array(lngEmailID)

  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset

  sSQL = "Select Name From ASRSysEmailGroupName Where EmailGroupID = " & lngEmailID
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If rsTemp.BOF And rsTemp.EOF Then
    GetEmailGroupName = vbNullString
  Else
    GetEmailGroupName = rsTemp(0)
  End If

  rsTemp.Close
  Set rsTemp = Nothing

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function


Public Function GetQuickEntryTables(plngScreenID As Long) As ADODB.Recordset
  ' Return a recordset of all the table id's of the controls which aren't related to the base table
  ' in the given screen.
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
    
  sSQL = "exec sp_ASRGetQuickEntryTables " & plngScreenID
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetQuickEntryTables = rsTemp
    
End Function

Public Function GetQuickEntryViews(plngTableID As Long) As ADODB.Recordset
  ' Return a recordset of the user defined views on the given table.
  Dim sSQL As String
  Dim rsViews As ADODB.Recordset
    
  sSQL = "SELECT viewID, viewName" & _
    " FROM ASRSysViews" & _
    " WHERE viewTableID = " & Trim(Str(plngTableID))
  Set rsViews = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetQuickEntryViews = rsViews

End Function

''Public Function GetViewPrivileges() As ADODB.Recordset
''
''    Dim sSQL As String
''    Dim rsTemp As ADODB.Recordset
''
''    sSQL = "exec sp_ASRGetViewPrivileges '" & gsUserGroup & "'"
''    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
''    Set GetViewPrivileges = rsTemp
''
''End Function

Public Function GetLookupTables() As ADODB.Recordset

    Dim sSQL As String
    Dim rsTables As ADODB.Recordset
    
    sSQL = "Select TableName, TableID From ASRSysTables Where TableType = 3"
    Set rsTables = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Set GetLookupTables = rsTables

End Function

Public Function GetLookupRecords(sTable As String, sColumn As String) As ADODB.Recordset

    Dim sSQL As String
    Dim rsRecords As ADODB.Recordset
    
    sSQL = "Select " & sColumn & " As Record, ID As RecordID From " & sTable
    Set rsRecords = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Set GetLookupRecords = rsRecords
    
End Function

Public Function GetColumns(lTableID As Long) As ADODB.Recordset

    Dim sSQL As String
    Dim rsColumns As ADODB.Recordset
    
    sSQL = "Select ColumnName, ColumnID From ASRSysColumns Where TableID = " & lTableID
    Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Set GetColumns = rsColumns

End Function


'Private Function CheckForLock(sUser As String, sLock As String) As Boolean
'
'    Dim sSQL As String
'    Dim rsTemp As ADODB.Recordset
'
'    sSQL = "exec sp_ASRGetLockInfo"
'    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'
'    If rsTemp(0) <> "No Lock" Then
'        If rsTemp(0) <> sUser Then
'            CheckForLock = True
'        End If
'        sLock = rsTemp(0)
'    End If
'
'    rsTemp.Close
'    Set rsTemp = Nothing
'
'End Function

Public Function RecordCount(sTable As String) As Long

    Dim sSQL As String
    Dim rsTemp As ADODB.Recordset
    
    sSQL = "SELECT COUNT(id) FROM " & sTable
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    RecordCount = rsTemp(0)
    
    rsTemp.Close
    Set rsTemp = Nothing
    
End Function




Public Function GetModuleArray(psModuleKey As String, psParameterKey As String) As ADODB.Recordset
  ' Return the recordset of the given module related columns.
  Dim sSQL As String
        
  sSQL = "SELECT sourceColumnID," & _
    " destColumnID" & _
    " FROM ASRSysModuleRelatedColumns" & _
    " WHERE moduleKey = '" & psModuleKey & "'" & _
    " AND parameterKey = '" & psParameterKey & "'"
  Set GetModuleArray = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
End Function
Public Function GetModuleParameter(psModuleKey As String, psParameterKey As String) As String
  ' Return the value of the given module parameter.
  Dim sSQL As String
  Dim rsModule As ADODB.Recordset
        
  sSQL = "SELECT parameterValue" & _
    " FROM ASRSysModuleSetup" & _
    " WHERE moduleKey = '" & psModuleKey & "'" & _
    " AND parameterKey = '" & psParameterKey & "'"
  Set rsModule = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
  If Not (rsModule.BOF And rsModule.EOF) Then
    If IsNull(rsModule!parameterValue) Then
      GetModuleParameter = vbNullString
    Else
      GetModuleParameter = rsModule!parameterValue
    End If
  Else
    GetModuleParameter = vbNullString
  End If
  
  Set rsModule = Nothing

End Function

Public Function GetFilteredIDs(plngExprID As Long) As String
' JPD - use the FilteredIDs function now.

'  ' Return a string describing the record IDs from the given table
'  ' that satisfay the given criteria.
'  Dim sIDs As String
'  Dim objExpr As clsExprExpression
'
'  Screen.MousePointer = vbHourglass
'
'  Set objExpr = New clsExprExpression
'  With objExpr
'    ' Initialise the filter expression object.
'    .ExpressionID = plngExprID
'    sIDs = .FilterCode
'  End With
'  Set objExpr = Nothing
'
'  Screen.MousePointer = vbDefault
'
'  GetFilteredIDs = sIDs

End Function

'Public Function TrainingBookingEnabled() As Boolean
'  ' Return TRUE if the Training Booking module has been enabled.
'  Dim lngCustNo As Long
'  Dim sSQL As String
'  Dim sAuth As String
'  Dim objLicence As ASRLicense.CLicense
'  Dim rsConfig As ADODB.Recordset
'
'  ' Get the Customer number and Module authorisation code.
'  sSQL = "SELECT * FROM ASRSysConfig"
'  Set rsConfig = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  If Not rsConfig.BOF And Not rsConfig.EOF Then
'    lngCustNo = IIf(IsNull(rsConfig!CustNo), 0, rsConfig!CustNo)
'    sAuth = IIf(IsNull(rsConfig!ModuleCode), "", rsConfig!ModuleCode)
'  End If
'  rsConfig.Close
'  Set rsConfig = Nothing
'
'  ' Validate the Training Booking module.
'  Set objLicence = New ASRLicense.CLicense
'  TrainingBookingEnabled = objLicence.GetModule(Training, sAuth, lngCustNo)
'  Set objLicence = Nothing
'
'End Function
'
'Public Function AbsenceEnabled() As Boolean
'  ' Return TRUE if the Absence module has been enabled.
'  Dim lngCustNo As Long
'  Dim sSQL As String
'  Dim sAuth As String
'  Dim objLicence As ASRLicense.CLicense
'  Dim rsConfig As ADODB.Recordset
'
'  ' Get the Customer number and Module authorisation code.
'  sSQL = "SELECT * FROM ASRSysConfig"
'  Set rsConfig = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  If Not rsConfig.BOF And Not rsConfig.EOF Then
'    lngCustNo = IIf(IsNull(rsConfig!CustNo), 0, rsConfig!CustNo)
'    sAuth = IIf(IsNull(rsConfig!ModuleCode), "", rsConfig!ModuleCode)
'  End If
'  rsConfig.Close
'  Set rsConfig = Nothing
'
'  ' Validate the Absence module.
'  Set objLicence = New ASRLicense.CLicense
'  AbsenceEnabled = objLicence.GetModule(Absence, sAuth, lngCustNo)
'  Set objLicence = Nothing
'
'End Function

Public Function BankHolidaysEnabled() As Boolean

  gfBankHolidaysEnabled = (glngBHolRegionTableID > 0) And _
                          (glngBHolRegionID > 0) And _
                          (glngBHolTableID > 0) And _
                          (glngBHolDateID > 0) And _
                          (glngBHolDateID > 0)
                          
End Function

'Public Function PersonnelEnabled() As Boolean
'  ' Return TRUE if the Personnel module has been enabled.
'  Dim lngCustNo As Long
'  Dim sSQL As String
'  Dim sAuth As String
'  Dim objLicence As ASRLicense.CLicense
'  Dim rsConfig As ADODB.Recordset
'
'  ' Get the Customer number and Module authorisation code.
'  sSQL = "SELECT * FROM ASRSysConfig"
'  Set rsConfig = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  If Not rsConfig.BOF And Not rsConfig.EOF Then
'    lngCustNo = IIf(IsNull(rsConfig!CustNo), 0, rsConfig!CustNo)
'    sAuth = IIf(IsNull(rsConfig!ModuleCode), "", rsConfig!ModuleCode)
'  End If
'  rsConfig.Close
'  Set rsConfig = Nothing
'
'  ' Validate the Personnel module.
'  Set objLicence = New ASRLicense.CLicense
'  PersonnelEnabled = objLicence.GetModule(Personnel, sAuth, lngCustNo)
'  Set objLicence = Nothing
'
'End Function

Public Function AFDEnabled() As AFDTypes

' Returns the type of AFD Module which has been enabled and installed.

  'Dim lngCustNo As Long
  'Dim sSQL As String
  'Dim sAuth As String
  'Dim objLicence As ASRLicense.CLicense
  'Dim rsConfig As ADODB.Recordset
  'Dim nnFlags As Long
  'Dim RecNo As Long
  'Dim Result As Long
  Dim bLicensed As Boolean
  
'  ' Get the Customer number and Module authorisation code.
'  sSQL = "SELECT * FROM ASRSysConfig"
'  Set rsConfig = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  If Not rsConfig.BOF And Not rsConfig.EOF Then
'    lngCustNo = IIf(IsNull(rsConfig!CustNo), 0, rsConfig!CustNo)
'    sAuth = IIf(IsNull(rsConfig!ModuleCode), "", rsConfig!ModuleCode)
'  End If
'  rsConfig.Close
'  Set rsConfig = Nothing
'
'  ' Validate the Afd module.
'  Set objLicence = New ASRLicense.CLicense
'  bLicensed = objLicence.GetModule(Afd, sAuth, lngCustNo)
'  Set objLicence = Nothing

  bLicensed = IsModuleEnabled(modAFD)

  ' If AFD is licensed, find out what version is installed on this machine
  AFDEnabled = IIf(bLicensed, modAFDSpecifics.AFDVersion, AFD_Disabled)


  If AFDEnabled = AFD_Disabled And bLicensed Then
    'MH20010830 New version stuff
    'If GetPCSetting( "Afd", "ShowError" & CStr(App.Major & "." & App.Minor & "." & App.Revision), True) = True Then
    If GetPCSetting("Afd", "ShowError" & CStr(app.Major & "." & app.Minor), True) = True Then
      If Not ASRDEVELOPMENT And Not gblnBatchJobsOnly Then
        COAMsgBox "OpenHR is configured to use Afd Postcode Names and Numbers software," & vbCrLf & _
               "but one or more of the required files are missing from your PC." & vbCrLf & vbCrLf & _
               "OpenHR will continue, but all Afd functionality will be disabled." & vbCrLf & vbCrLf & _
               "(This message will not appear again until a new version of OpenHR is installed.)" _
               , vbExclamation + vbOKOnly, Application.Name
        'MH20010830 New version stuff
        'SavePCSetting "Afd", "ShowError" & CStr(App.Major & "." & App.Minor & "." & App.Revision), False
        SavePCSetting "Afd", "ShowError" & CStr(app.Major & "." & app.Minor), False
      End If
    End If
  End If
  
End Function

Public Function EmailGenerationEnabled()

  Dim objTestConn As ADODB.Connection
  Dim adoCmd As ADODB.Command
  Dim sSQL As String
  Dim strOutput As String
  Dim lngCount As Long
  Dim strTestEmailAddr As String
  Dim lngMethod As Long
  Dim sMessage As String

  On Error GoTo LocalErr

  strTestEmailAddr = GetSystemSetting("Email", "Test Messages", "")
  strOutput = vbNullString
    
  If strTestEmailAddr <> vbNullString Then

    Set objTestConn = New ADODB.Connection
    With objTestConn
      'JPD 20070529 Fault 12220
      '.ConnectionString = gADOCon.ConnectionString
      '.ConnectionString = GetConnectionString("APP=", "Test Emails") & IIf(gbUseWindowsAuthentication, ";Integrated Security=SSPI;", "")
      .ConnectionString = GetConnectionString("Application Name=", "Test Emails") & IIf(gbUseWindowsAuthentication, ";Integrated Security=SSPI;", "")
      .Provider = "SQLOLEDB"
      .CommandTimeout = 30
      .ConnectionTimeout = 30
      .CursorLocation = adUseServer
      .Mode = adModeReadWrite
      .Properties("Packet Size") = 32767
      .Open
    End With
  
    objTestConn.Errors.Clear
   
    lngMethod = GetSystemSetting("Email", "Method", 1)
    Select Case lngMethod
    Case 1
      sSQL = "IF (SELECT Count(*) FROM AsrSysEmailLinks) > 0 " & vbCrLf & _
             "  exec master.dbo.xp_startmail"
'    Case 2
'      sSQL = "IF (SELECT Count(*) FROM AsrSysEmailLinks) > 0 " & vbCrLf & _
'             "  exec msdb.dbo.sysmail_start_sp"
    Case Else
      sSQL = vbNullString
    End Select


    If sSQL <> vbNullString Then
      
      Set adoCmd = New ADODB.Command
      adoCmd.ActiveConnection = objTestConn
      adoCmd.CommandTimeout = 30
      adoCmd.CommandText = sSQL
      adoCmd.Execute
      
      For lngCount = 0 To objTestConn.Errors.Count - 1
        If Trim(objTestConn.Errors(lngCount).Description) <> vbNullString Then
        
          If UCase(objTestConn.Errors(lngCount).Description) <> "SQL MAIL SESSION STARTED." And _
             UCase(objTestConn.Errors(lngCount).Description) <> "SQL MAIL SESSION IS ALREADY STARTED." Then
            strOutput = strOutput & objTestConn.Errors(lngCount).Description & vbCrLf
          End If
    
        End If
      Next
      Set adoCmd = Nothing
    End If
  
    objTestConn.Errors.Clear
    Set adoCmd = New ADODB.Command
    adoCmd.ActiveConnection = objTestConn
    adoCmd.CommandTimeout = 30
    adoCmd.CommandText = "IF (SELECT Count(*) FROM AsrSysEmailLinks) > 0 " & _
                         "BEGIN" & vbCrLf & _
                         "  EXEC spASRSendMail 0, '" & Replace(strTestEmailAddr, "'", "''") & "','','','','" & _
                         Replace(gsUserName & vbCrLf & gsDatabaseName & vbCrLf & UI.GetHostName, "'", "''") & "',''" & vbCrLf & _
                         "END"
    adoCmd.Execute
  
    For lngCount = 0 To objTestConn.Errors.Count - 1
    
      sMessage = UCase(Trim(objTestConn.Errors(lngCount).Description))
    
      If sMessage <> vbNullString Then
        
        If sMessage <> "MAIL SENT." And sMessage <> "MAIL QUEUED." And _
            Not (InStr(1, sMessage, "MAIL (ID:") > 0 And InStr(1, sMessage, ") QUEUED.") > 0) Then
          strOutput = strOutput & sMessage & vbCrLf
        End If
      
      End If
    Next
  End If

  Set adoCmd = Nothing
  If Not objTestConn Is Nothing Then
    If objTestConn.State = adStateOpen Then
      objTestConn.Close
    End If
    Set objTestConn = Nothing
  End If

LocalErr:
  
  If Err.Number <> 0 Then
    strOutput = Err.Description
  End If

  If strOutput <> vbNullString And Not gblnBatchJobsOnly Then
    COAMsgBox "OpenHR is configured to send emails dependant upon data entered " & _
           "but there is currently a problem sending emails from the SQL server." & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "OpenHR will continue, but the email link functionality will be disabled.", vbExclamation + vbOKOnly, Application.Name
  End If
  
  Set adoCmd = Nothing
  If Not objTestConn Is Nothing Then
    If objTestConn.State = adStateOpen Then
      objTestConn.Close
    End If
    Set objTestConn = Nothing
  End If

End Function



Public Function GetFirstOrderColumnName(plngOrderID As Long) As String
  ' Return the name of the first column in the given order.
  Dim sSQL As String
  Dim rsOrder As ADODB.Recordset
    
  sSQL = "SELECT columnName" & _
    " FROM ASRSysColumns" & _
    " JOIN ASRSysOrderItems ON ASRSysOrderItems.columnID = ASRSysColumns.columnID" & _
    " WHERE ASRSysOrderItems.orderID = " & Trim(Str(plngOrderID)) & _
    " AND type = 'O'" & _
    " AND ASRSysOrderItems.sequence = 1"
  Set rsOrder = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If rsOrder.BOF And rsOrder.EOF Then
    GetFirstOrderColumnName = ""
  Else
    If IsNull(rsOrder!ColumnName) Then
      GetFirstOrderColumnName = ""
    Else
      GetFirstOrderColumnName = rsOrder!ColumnName
    End If
  End If
  
  rsOrder.Close
  Set rsOrder = Nothing
  
End Function



Public Function GetPrimaryTableSubMenu(plngTableID As Long, psTableName As String) As Variant
  ' Return an array of information that can be used to format the given table's sub-menu
  ' on the Database menu for the current user.
  ' The array contains a row for each screen and view screen.
  '
  ' For each primary table the following information is given :
  ' screenID    ID of the screen
  ' screenName  Name of the screen
  ' viewID      ID of the view
  ' viewName    Name of the view
  ' screenPictureID ID of the screen's icon
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim sViewList As String
  Dim rsScreens As ADODB.Recordset
  Dim avScreenInfo() As Variant
  Dim objTableView As CTablePrivilege
  
  ' Create an array with records for each screen for each permitted view on the primary table.
  ReDim avScreenInfo(5, 0)
  
  ' Create a list of the current user's permitted views on the current table.
  sViewList = "0"
  For Each objTableView In gcoTablePrivileges.Collection
    If Not (objTableView.IsTable) And _
      (objTableView.TableID = plngTableID) And _
      (objTableView.AllowSelect) Then
      
      sViewList = sViewList & ", " & Trim(Str(objTableView.ViewID))
    End If
  Next objTableView
  Set objTableView = Nothing

  sSQL = "SELECT ASRSysViewScreens.screenID," & _
    " ASRSysScreens.name," & _
    " ASRSysViews.viewID," & _
    " ASRSysViews.ViewName," & _
    " ASRSysScreens.PictureID" & _
    " FROM ASRSysViews" & _
    " INNER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID" & _
    " INNER JOIN ASRSysScreens ON ASRSysViewScreens.screenID = ASRSysScreens.screenID" & _
    " WHERE ASRSysViews.viewID IN (" & sViewList & ")"
  
  If gcoTablePrivileges.Item(psTableName).AllowSelect Then
    ' The current user does have SELECT permission on the given table, so populate the array
    ' table with records for each screen associated with the primary table.
    sSQL = sSQL & _
      " UNION" & _
      " SELECT ASRSysScreens.screenID," & _
      " ASRSysScreens.Name," & _
      " 0 AS viewID," & _
      " '' AS viewName," & _
      " ASRSysScreens.pictureID" & _
      " FROM ASRSysScreens" & _
      " WHERE (ASRSysScreens.tableID = " & Trim(Str(plngTableID)) & ")" & _
      " AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))" & _
      " AND ((ASRSysScreens.quickEntry IS null) OR (ASRSysScreens.quickEntry = 0))"
  End If
  
  sSQL = sSQL & _
    " ORDER BY ASRSysScreens.name DESC, viewName DESC"

  Set rsScreens = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Do While Not rsScreens.EOF
    iNextIndex = UBound(avScreenInfo, 2) + 1
    ReDim Preserve avScreenInfo(5, iNextIndex)
    avScreenInfo(1, iNextIndex) = rsScreens!ScreenID
    avScreenInfo(2, iNextIndex) = rsScreens!Name
    avScreenInfo(3, iNextIndex) = rsScreens!ViewID
    avScreenInfo(4, iNextIndex) = rsScreens!ViewName
    avScreenInfo(5, iNextIndex) = rsScreens!PictureID
    
    rsScreens.MoveNext
  Loop
  rsScreens.Close
  Set rsScreens = Nothing

  GetPrimaryTableSubMenu = avScreenInfo
  
End Function

Public Function SystemPermission(psCategoryKey As String, psPermissionKey As String) As Boolean
  
  On Error GoTo SettingNotFound
  
  If gfCurrentUserIsSysSecMgr Then
    SystemPermission = True
    GoTo TidyUpAndExit
  End If
  
  ' If value is not found error is raised and stored procedure fires
  If gcolSystemPermissions.Item(psCategoryKey & "," & psPermissionKey) Then
    SystemPermission = gcolSystemPermissions.Item(psCategoryKey & "," & psPermissionKey)
  End If

TidyUpAndExit:
  Exit Function

SettingNotFound:

  On Error GoTo ErrorTrap

    ' Return TRUE if the user has been granted the given permission.
    Dim cmADO As ADODB.Command
    Dim pmADO As ADODB.Parameter
  
    ' Check if the user can create New instances of the given category.
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "sp_ASRSystemPermission"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
  
      Set pmADO = .CreateParameter("permissionGranted", adBoolean, adParamOutput)
      .Parameters.Append pmADO
      
      Set pmADO = .CreateParameter("categoryKey", adVarChar, adParamInput, 50)
      .Parameters.Append pmADO
      pmADO.Value = psCategoryKey
  
      Set pmADO = .CreateParameter("permissionKey", adVarChar, adParamInput, 50)
      .Parameters.Append pmADO
      pmADO.Value = psPermissionKey
  
      Set pmADO = .CreateParameter("LoggedInSQLUser", adVarChar, adParamInput, 50)
      .Parameters.Append pmADO
      pmADO.Value = gsSQLUserName
  
      Set pmADO = Nothing
  
      cmADO.Execute
  
      SystemPermission = .Parameters(0).Value
    End With
    Set cmADO = Nothing

    ' Adds this setting to the permissions object
    gcolSystemPermissions.Add SystemPermission, psCategoryKey & "," & psPermissionKey
    
    Exit Function

ErrorTrap:

  If InStr(1, Err.Description, CONNECTIONBROKEN_MESSAGE, vbTextCompare) Then
    datGeneral.ReEstablishADOConnection
    Resume SettingNotFound
  End If


End Function


Public Function UniqueSQLObjectName(strPrefix As String, intType As Integer) As String

'  Dim rsTemp As ADODB.Recordset
'  Dim strSQL As String
'
'  strSQL = "sp_ASRUniqueObjectName '" & strPrefix & "', " & CStr(intType)
'  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
'
'  UniqueSQLObjectName = rsTemp!UniqueObject
'
'  Set rsTemp = Nothing
  
  'TM20020530 Fault 3756 - function altered as the sp needs to insert a record into a table
  'before returning a value, so collect the returned parameter rather than a recordset.
  
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
    pmADO.Value = strPrefix
    
    Set pmADO = .CreateParameter("Type", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = intType
    
    Set pmADO = Nothing
    
    .Execute
    
    UniqueSQLObjectName = IIf(IsNull(.Parameters(0).Value), vbNullString, .Parameters(0).Value)
      
  End With

  Set cmdUniqObj = Nothing
  
  
End Function

Public Function DropUniqueSQLObject(sSQLObjectName As String, iType As Integer) As Boolean

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.DropUniqueSQLObject(sSQLObjectName, iType)", Array(sSQLObjectName, iType)

  Dim cmdUniqObj As New ADODB.Command
  Dim pmADO As ADODB.Parameter
 
  If Len(sSQLObjectName) > 0 Then
    With cmdUniqObj
      .CommandText = "sp_ASRDropUniqueObject"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
                
      Set pmADO = .CreateParameter("UniqueObjectName", adLongVarChar, adParamInput, -1)
      .Parameters.Append pmADO
      pmADO.Value = sSQLObjectName
      
      Set pmADO = .CreateParameter("Type", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = iType
      
      Set pmADO = Nothing
      
      .Execute
    End With
  End If
  
  DropUniqueSQLObject = True
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Set cmdUniqObj = Nothing
  Exit Function
ErrorTrap:
  DropUniqueSQLObject = False
  gobjErrorStack.HandleError

End Function

Public Function IsAChildOf(lTestTableID As Long, lBaseTableID As Long) As Boolean

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String

  strSQL = "SELECT * FROM ASRSysRelations WHERE ParentID = " & lBaseTableID & _
  " AND ChildID = " & lTestTableID
  
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  If rsTemp.BOF And rsTemp.EOF Then
    IsAChildOf = False
  Else
    IsAChildOf = True
  End If
  
  Set rsTemp = Nothing

End Function

Public Function IsAParentOf(lTestTableID As Long, lBaseTableID As Long) As Boolean

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String

  strSQL = "SELECT * FROM ASRSysRelations WHERE ChildID = " & lBaseTableID & _
  " AND ParentID = " & lTestTableID
  
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  If rsTemp.BOF And rsTemp.EOF Then
    IsAParentOf = False
  Else
    IsAParentOf = True
  End If
  
  Set rsTemp = Nothing

End Function


Public Function CheckCanMakeHidden(pstrType As String, _
                                   plngID As Long, _
                                   pstrUser As String, _
                                   pstrCaption As String) As Boolean
                                   
  ' Recordset and its source string
  'Dim rsTemp As ADODB.Recordset
  Dim sSQL As String
  
  ' Count and util names of utils which will be made hidden if allowed
  Dim iCount_Owner As Integer
  Dim sDetails_Owner As String
  
  ' count and util names of utils which prevent the change if applicable
  Dim iCount_NotOwner As Integer
  Dim sDetails_NotOwner As String
  
  ' comma separated list of the utility IDs
  Dim sCrossTabIDs As String
  Dim sCustomReportsIDs As String
  Dim sCalendarReportsIDs As String
  Dim sDataTransferIDs As String
  Dim sExportIDs As String
  Dim sGlobalAddIDs As String
  Dim sGlobalUpdateIDs As String
  Dim sGlobalDeleteIDs As String
  Dim sMailMergeIDs As String
  Dim sLabelIDs As String
  Dim sRecordProfileIDs As String
  Dim sMatchReportIDs As String
  Dim sSuccessionPlanningIDs As String
  Dim sCareerProgressionIDs As String
  
  ' batch job info in which utils which require changing are contained
  Dim sBatchJobDetails_Owner As String
  Dim sBatchJobDetails_NotOwner As String
  Dim sBatchJobDetails_ScheduledForOtherUsers As String
  Dim sBatchJobIDs As String
  Dim fBatchJobsOK As Boolean
  Dim sBatchJobScheduledUserGroups As String
  
  Dim sExprIDs As String
  Dim sCalculationIDs As String
  Dim sFilterIDs As String
  Dim sSuperExprIDs As String
  
  fBatchJobsOK = True
  
  'force the username to lowercase for comparisons in the function.
  pstrUser = LCase(pstrUser)
  
  Select Case UCase(pstrType)
  
    '*****************************************************
    ' Calculations/Filters
    '*****************************************************
    Case "E", "F"
      '---------------------------------------------------
      ' Check Calculations/Filters For This Expression
      ' NB. This check must be made before checking the reports/utilities
      '---------------------------------------------------
      sExprIDs = CStr(plngID)
      sSuperExprIDs = GetAllExprRootIDs(plngID)
      
      If sSuperExprIDs <> vbNullString Then
        sExprIDs = sExprIDs & "," & sSuperExprIDs
        
        sSQL = "SELECT ASRSysExpressions.Name," & _
          "   ASRSysExpressions.exprID AS [ID]," & _
          "   ASRSysExpressions.Username," & _
          "   ASRSysExpressions.Access" & _
          " FROM ASRSysExpressions" & _
          " WHERE ASRSysExpressions.exprID IN (" & sSuperExprIDs & ")" & _
          "   AND ASRSysExpressions.type = " & giEXPR_RUNTIMECALCULATION
      
        CheckForPicklistsExpressions utlCalculation, _
          sSQL, _
          pstrUser, _
          iCount_Owner, _
          sDetails_Owner, _
          sCalculationIDs, _
          iCount_NotOwner, _
          sDetails_NotOwner
      
        sSQL = "SELECT ASRSysExpressions.Name," & _
          "   ASRSysExpressions.exprID AS [ID]," & _
          "   ASRSysExpressions.Username," & _
          "   ASRSysExpressions.Access" & _
          " FROM ASRSysExpressions" & _
          " WHERE ASRSysExpressions.exprID IN (" & sSuperExprIDs & ")" & _
          "   AND ASRSysExpressions.type = " & giEXPR_RUNTIMEFILTER
      
        CheckForPicklistsExpressions utlFilter, _
          sSQL, _
          pstrUser, _
          iCount_Owner, _
          sDetails_Owner, _
          sFilterIDs, _
          iCount_NotOwner, _
          sDetails_NotOwner
      End If
  
      '---------------------------------------------------
      ' Check Calendar Reports For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysCalendarReports.Name," & _
        "   AsrSysCalendarReports.ID," & _
        "   AsrSysCalendarReports.Username," & _
        "   COUNT (ASRSYSCalendarReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysCalendarReports" & _
        " LEFT OUTER JOIN ASRSYSCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSYSCalendarReportEvents.calendarReportID " & _
        " LEFT OUTER JOIN ASRSYSCalendarReportAccess ON AsrSysCalendarReports.ID = ASRSYSCalendarReportAccess.ID" & _
        "   AND ASRSYSCalendarReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCalendarReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)"
      sSQL = sSQL & _
        " WHERE AsrSysCalendarReports.DescriptionExpr IN (" & sExprIDs & ")" & _
        "   OR AsrSysCalendarReports.StartDateExpr IN (" & sExprIDs & ")" & _
        "   OR AsrSysCalendarReports.EndDateExpr IN (" & sExprIDs & ")" & _
        "   OR ASRSysCalendarReports.Filter IN (" & sExprIDs & ")" & _
        "   OR ASRSYSCalendarReportEvents.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY AsrSysCalendarReports.Name," & _
        "   AsrSysCalendarReports.ID," & _
        "   AsrSysCalendarReports.Username"
      CheckForPicklistsExpressions utlCalendarReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCalendarReportsIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these Calendar Reports are contained within a batch job
      If Len(Trim(sCalendarReportsIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCalendarReport, _
          sCalendarReportsIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Career Progression For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 2 " & _
        "  AND (ASRSysMatchReportName.table1Filter IN (" & sExprIDs & ")" & _
        "  OR ASRSysMatchReportName.table2Filter IN (" & sExprIDs & "))" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlCareer, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCareerProgressionIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sCareerProgressionIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCareer, _
          sCareerProgressionIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Cross Tabs For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysCrossTab.Name," & _
        "   AsrSysCrossTab.[CrossTabID] AS [ID]," & _
        "   AsrSysCrossTab.Username," & _
        "   COUNT (ASRSYSCrossTabAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysCrossTab" & _
        " LEFT OUTER JOIN ASRSYSCrossTabAccess ON AsrSysCrossTab.crossTabID = ASRSYSCrossTabAccess.ID" & _
        "   AND ASRSYSCrossTabAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCrossTabAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysCrossTab.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY AsrSysCrossTab.Name," & _
        "   AsrSysCrossTab.crossTabID," & _
        "   AsrSysCrossTab.Username"
      CheckForPicklistsExpressions utlCrossTab, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCrossTabIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
    
      ' Now check that any of these CrossTabs are contained within a batch job
      If Len(Trim(sCrossTabIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCrossTab, _
          sCrossTabIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Custom Reports For This Expression
      '---------------------------------------------------
      sSQL = "SELECT ASRSysCustomReportsName.Name," & _
        "   ASRSysCustomReportsName.ID," & _
        "   ASRSysCustomReportsName.Username," & _
        "   COUNT (ASRSYSCustomReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSysCustomReportsDetails ON ASRSysCustomReportsName.ID = AsrSysCustomReportsDetails.CustomReportID" & _
        " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
        " LEFT OUTER JOIN ASRSYSCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSYSCustomReportAccess.ID" & _
        "   AND ASRSYSCustomReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCustomReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)"
      sSQL = sSQL & _
        " WHERE ASRSysCustomReportsName.Filter IN (" & sExprIDs & ")" & _
        "   OR ASRSysCustomReportsName.Parent1Filter IN (" & sExprIDs & ")" & _
        "   OR ASRSysCustomReportsName.Parent2Filter IN (" & sExprIDs & ")" & _
        "   OR ASRSYSCustomReportsChildDetails.ChildFilter IN (" & sExprIDs & ")" & _
        "   OR(AsrSysCustomReportsDetails.Type = 'E' " & _
        "     AND AsrSysCustomReportsDetails.ColExprID IN (" & sExprIDs & "))" & _
        " GROUP BY ASRSysCustomReportsName.Name," & _
        "   ASRSysCustomReportsName.ID," & _
        "   ASRSysCustomReportsName.Username"
      
      CheckForPicklistsExpressions utlCustomReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCustomReportsIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
          
      ' Now check that any of these Custom Reports are contained within a batch job
      If Len(Trim(sCustomReportsIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCustomReport, _
          sCustomReportsIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Data Transfer For This Filter
      '---------------------------------------------------
      sSQL = "SELECT AsrSysDataTransferName.Name," & _
        "   AsrSysDataTransferName.DataTransferID AS [ID]," & _
        "   AsrSysDataTransferName.Username," & _
        "   COUNT (ASRSYSDataTransferAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysDataTransferName" & _
        " LEFT OUTER JOIN ASRSYSDataTransferAccess ON AsrSysDataTransferName.DataTransferID = ASRSYSDataTransferAccess.ID" & _
        "   AND ASRSYSDataTransferAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSDataTransferAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysDataTransferName.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY AsrSysDataTransferName.Name," & _
        "   AsrSysDataTransferName.DataTransferID," & _
        "   AsrSysDataTransferName.Username"
      CheckForPicklistsExpressions utlDataTransfer, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sDataTransferIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
                    
      ' Now check that any of these DataTransfers are contained within a batch job
      If Len(Trim(sDataTransferIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlDataTransfer, _
          sDataTransferIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Envelopes & Labels For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID AS [ID]," & _
        "   AsrSysMailMergeName.Username," & _
        "   COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysMailMergeName" & _
        " LEFT OUTER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID" & _
        " LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.mailMergeID = ASRSYSMailMergeAccess.ID" & _
        "   AND ASRSYSMailMergeAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)"
      sSQL = sSQL & _
        " WHERE AsrSysMailMergeName.isLabel = 1" & _
        " AND ((AsrSysMailMergeName.FilterID IN (" & sExprIDs & "))" & _
        "   OR (AsrSysMailMergeColumns.Type = 'E' " & _
        "     AND AsrSysMailMergeColumns.ColumnID IN (" & sExprIDs & ")))" & _
        " GROUP BY AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID," & _
        "   AsrSysMailMergeName.Username"
      CheckForPicklistsExpressions utlLabel, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sLabelIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
        
      ' Now check if any of these Envelopes & Labels are contained within a batch job
      If Len(Trim(sLabelIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlLabel, _
          sLabelIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Export For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysExportName.Name," & _
        "   AsrSysExportName.ID," & _
        "   AsrSysExportName.Username," & _
        "   COUNT (ASRSYSExportAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysExportName" & _
        " LEFT OUTER JOIN AsrSysExportDetails ON AsrSysExportName.ID = AsrSysExportDetails.exportID" & _
        " LEFT OUTER JOIN ASRSYSExportAccess ON AsrSysExportName.ID = ASRSYSExportAccess.ID" & _
        "   AND ASRSYSExportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSExportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)"
      sSQL = sSQL & _
        " WHERE AsrSysExportName.Filter IN (" & sExprIDs & ")" & _
        "   OR AsrSysExportName.Parent1Filter IN (" & sExprIDs & ")" & _
        "   OR AsrSysExportName.Parent2Filter IN (" & sExprIDs & ")" & _
        "   OR AsrSysExportName.ChildFilter IN (" & sExprIDs & ")" & _
        "   OR (AsrSysExportDetails.Type = 'X' " & _
        "     AND AsrSysExportDetails.ColExprID IN (" & sExprIDs & "))" & _
        " GROUP BY AsrSysExportName.Name," & _
        "   AsrSysExportName.ID," & _
        "   AsrSysExportName.Username"
      CheckForPicklistsExpressions utlExport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sExportIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
            
      ' Now check that any of these Exports are contained within a batch job
      If Len(Trim(sExportIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlExport, _
          sExportIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Global Add For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN AsrSysGlobalItems ON AsrSysGlobalFunctions.functionID = AsrSysGlobalItems.FunctionID" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)"
      sSQL = sSQL & _
        " WHERE AsrSysGlobalFunctions.Type = 'A' " & _
        "   AND ((AsrSysGlobalFunctions.FilterID IN (" & sExprIDs & "))" & _
        "     OR (AsrSysGlobalItems.ValueType = 4 " & _
        "     AND AsrSysGlobalItems.ExprID IN (" & sExprIDs & ")))" & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      CheckForPicklistsExpressions UtlGlobalAdd, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalAddIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these Global Adds are contained within a batch job
      If Len(Trim(sGlobalAddIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs UtlGlobalAdd, _
          sGlobalAddIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Global Update For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN AsrSysGlobalItems ON AsrSysGlobalFunctions.functionID = AsrSysGlobalItems.FunctionID" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)"
      sSQL = sSQL & _
        " WHERE AsrSysGlobalFunctions.Type = 'U' " & _
        "  AND ((AsrSysGlobalFunctions.FilterID IN (" & sExprIDs & "))" & _
        "    OR (AsrSysGlobalItems.ValueType = 4 " & _
        "      AND AsrSysGlobalItems.ExprID IN (" & sExprIDs & ")))" & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      CheckForPicklistsExpressions utlGlobalUpdate, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalUpdateIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
            
      ' Now check that any of these Global Updates are contained within a batch job
      If Len(Trim(sGlobalUpdateIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlGlobalUpdate, _
          sGlobalUpdateIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Global Delete For This Filter
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysGlobalFunctions.Type = 'D' " & _
        "  AND AsrSysGlobalFunctions.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      CheckForPicklistsExpressions utlGlobalDelete, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalDeleteIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these Global Deletes are contained within a batch job
      If Len(Trim(sGlobalDeleteIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlGlobalDelete, _
          sGlobalDeleteIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Mail Merge For This Expression
      '---------------------------------------------------
      sSQL = "SELECT AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID AS [ID]," & _
        "   AsrSysMailMergeName.Username," & _
        "   COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysMailMergeName" & _
        " LEFT OUTER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID" & _
        " LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.mailMergeID = ASRSYSMailMergeAccess.ID" & _
        "   AND ASRSYSMailMergeAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)"
      sSQL = sSQL & _
        " WHERE AsrSysMailMergeName.isLabel = 0" & _
        "   AND ((AsrSysMailMergeName.FilterID IN (" & sExprIDs & "))" & _
        "     OR (AsrSysMailMergeColumns.Type = 'E' " & _
        "       AND AsrSysMailMergeColumns.ColumnID IN (" & sExprIDs & ")))" & _
        " GROUP BY AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID," & _
        "   AsrSysMailMergeName.Username"
      CheckForPicklistsExpressions utlMailMerge, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sMailMergeIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
        
      ' Now check if any of these Merges are contained within a batch job
      If Len(Trim(sMailMergeIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlMailMerge, _
          sMailMergeIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Match Report For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 0 " & _
        "  AND (ASRSysMatchReportName.table1Filter IN (" & sExprIDs & ")" & _
        "  OR ASRSysMatchReportName.table2Filter IN (" & sExprIDs & "))" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlMatchReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sMatchReportIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sMatchReportIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlMatchReport, _
          sMatchReportIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Record Profiles For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysRecordProfileName.Name," & _
        "   ASRSysRecordProfileName.recordProfileID AS [ID]," & _
        "   ASRSysRecordProfileName.Username," & _
        "   COUNT (ASRSYSRecordProfileAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysRecordProfileName" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileAccess.ID" & _
        "   AND ASRSYSRecordProfileAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSRecordProfileAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysRecordProfileName.FilterID IN (" & sExprIDs & ")" & _
        "   OR ASRSYSRecordProfileTables.FilterID IN (" & sExprIDs & ")" & _
        " GROUP BY ASRSysRecordProfileName.Name," & _
        "   ASRSysRecordProfileName.recordProfileID," & _
        "   ASRSysRecordProfileName.Username"
      CheckForPicklistsExpressions utlRecordProfile, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sRecordProfileIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Record Profiles are contained within a batch job
      If Len(Trim(sRecordProfileIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlRecordProfile, _
          sRecordProfileIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Succession Planning For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 1 " & _
        "  AND (ASRSysMatchReportName.table1Filter IN (" & sExprIDs & ")" & _
        "  OR ASRSysMatchReportName.table2Filter IN (" & sExprIDs & "))" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      
      CheckForPicklistsExpressions utlSuccession, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sSuccessionPlanningIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sSuccessionPlanningIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlSuccession, _
          sSuccessionPlanningIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Ok, all relevant utility definitions have now been checked, so check
      ' the counts and act accordingly
      '---------------------------------------------------
      If (iCount_Owner = 0) And _
        (iCount_NotOwner = 0) And _
        fBatchJobsOK And _
        (Len(sBatchJobDetails_Owner) = 0) Then
          
        CheckCanMakeHidden = True
        Exit Function
      
      ElseIf (iCount_Owner > 0) And _
        (iCount_NotOwner = 0) And _
        fBatchJobsOK Then
        ' Can change utils and no utils
        ' are contained within batch jobs
        ' that cant be changed
'        If COAMsgBox("Changing the selected " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " to hidden will automatically" & vbCrLf & _
'                  "make the following definition(s), of which you are the" & vbCrLf & _
'                  "owner, hidden also:" & vbCrLf & vbCrLf & _
'                  sDetails_Owner & sBatchJobDetails_Owner & vbCrLf & _
'                  "Do you wish to continue ?", vbQuestion + vbYesNo, IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)) _
'                  = vbNo Then
'          Screen.MousePointer = vbDefault
'          CheckCanMakeHidden = False
'          Exit Function
        'NHRD12112004 Fault 6335
        If COAMsgBox("Changing the selected " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " to hidden will automatically " & _
                  "make the following definition(s)," & vbCrLf & " of which you are the owner, hidden also:" & vbCrLf & vbCrLf & _
                  sDetails_Owner & sBatchJobDetails_Owner & vbCrLf & _
                  "Do you wish to continue ?", vbQuestion + vbYesNo, IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)) _
                  = vbNo Then
          Screen.MousePointer = vbDefault
          CheckCanMakeHidden = False
          Exit Function
        Else
          ' Ok, we are continuing, so lets update all those utils to hidden !
          
          ' Calculations
          If Len(Trim(sCalculationIDs)) > 0 Then
            HideUtilities utlCalculation, sCalculationIDs
            Call UtilUpdateLastSavedMultiple(utlCalculation, sCalculationIDs)
          End If
                    
          ' Filters
          If Len(Trim(sFilterIDs)) > 0 Then
            HideUtilities utlFilter, sFilterIDs
            Call UtilUpdateLastSavedMultiple(utlFilter, sFilterIDs)
          End If
          
          ' Batch Jobs
          If Len(Trim(sBatchJobIDs)) > 0 Then
            HideUtilities utlBatchJob, sBatchJobIDs
            Call UtilUpdateLastSavedMultiple(utlBatchJob, sBatchJobIDs)
          End If
          
          ' Calendar Reports
          If Len(Trim(sCalendarReportsIDs)) > 0 Then
            HideUtilities utlCalendarReport, sCalendarReportsIDs
            Call UtilUpdateLastSavedMultiple(utlCalendarReport, sCalendarReportsIDs)
          End If
          
          ' Career Progression
          If Len(Trim(sCareerProgressionIDs)) > 0 Then
            HideUtilities utlCareer, sCareerProgressionIDs
            Call UtilUpdateLastSavedMultiple(utlCareer, sCareerProgressionIDs)
          End If
          
          ' Cross Tabs
          If Len(Trim(sCrossTabIDs)) > 0 Then
            HideUtilities utlCrossTab, sCrossTabIDs
            Call UtilUpdateLastSavedMultiple(utlCrossTab, sCrossTabIDs)
          End If
          
          ' Custom Reports
          If Len(Trim(sCustomReportsIDs)) > 0 Then
            HideUtilities utlCustomReport, sCustomReportsIDs
            Call UtilUpdateLastSavedMultiple(utlCustomReport, sCustomReportsIDs)
          End If
                   
          ' Data Transfer
          If Len(Trim(sDataTransferIDs)) > 0 Then
            HideUtilities utlDataTransfer, sDataTransferIDs
            Call UtilUpdateLastSavedMultiple(utlDataTransfer, sDataTransferIDs)
          End If
          
          ' Envelopes & Labels
          If Len(Trim(sLabelIDs)) > 0 Then
            HideUtilities utlLabel, sLabelIDs
            Call UtilUpdateLastSavedMultiple(utlLabel, sLabelIDs)
          End If
          
          ' Export
          If Len(Trim(sExportIDs)) > 0 Then
            HideUtilities utlExport, sExportIDs
            Call UtilUpdateLastSavedMultiple(utlExport, sExportIDs)
          End If
                             
          ' Global Add
          If Len(Trim(sGlobalAddIDs)) > 0 Then
            HideUtilities UtlGlobalAdd, sGlobalAddIDs
            Call UtilUpdateLastSavedMultiple(UtlGlobalAdd, sGlobalAddIDs)
          End If
          
          ' Global Update
          If Len(Trim(sGlobalUpdateIDs)) > 0 Then
            HideUtilities utlGlobalUpdate, sGlobalUpdateIDs
            Call UtilUpdateLastSavedMultiple(utlGlobalUpdate, sGlobalUpdateIDs)
          End If
          
          ' Global Delete
          If Len(Trim(sGlobalDeleteIDs)) > 0 Then
            HideUtilities utlGlobalDelete, sGlobalDeleteIDs
            Call UtilUpdateLastSavedMultiple(utlGlobalDelete, sGlobalDeleteIDs)
          End If
          
          ' Mail Merge
          If Len(Trim(sMailMergeIDs)) > 0 Then
            HideUtilities utlMailMerge, sMailMergeIDs
            Call UtilUpdateLastSavedMultiple(utlMailMerge, sMailMergeIDs)
          End If
          
          ' Match Reports
          If Len(Trim(sMatchReportIDs)) > 0 Then
            HideUtilities utlMatchReport, sMatchReportIDs
            Call UtilUpdateLastSavedMultiple(utlMatchReport, sMatchReportIDs)
          End If
          
          ' Record Profile
          If Len(Trim(sRecordProfileIDs)) > 0 Then
            HideUtilities utlRecordProfile, sRecordProfileIDs
            Call UtilUpdateLastSavedMultiple(utlRecordProfile, sRecordProfileIDs)
          End If
          
          ' Succession Planning
          If Len(Trim(sSuccessionPlanningIDs)) > 0 Then
            HideUtilities utlSuccession, sSuccessionPlanningIDs
            Call UtilUpdateLastSavedMultiple(utlSuccession, sSuccessionPlanningIDs)
          End If
          
          ' Ok, all done, so exit now
          CheckCanMakeHidden = True
          Exit Function
        
        End If
      
      ElseIf (iCount_Owner > 0) And _
        (iCount_NotOwner = 0) And _
        (Not fBatchJobsOK) Then
        ' Can change utils but abort cos those
        ' utils are in batch jobs which cannot
        ' be changed
        If Len(sBatchJobDetails_ScheduledForOtherUsers) > 0 Then
          COAMsgBox "This " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " cannot be made hidden as it is used in " & vbCrLf & _
                 "definition(s) which are included in the following" & vbCrLf & _
                 "batch jobs which are scheduled to be run by other user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, vbExclamation + vbOKOnly _
                 , IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)
        Else
          COAMsgBox "This " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " cannot be made hidden as it is used in " & vbCrLf & _
                 "definition(s) which are included in the following" & vbCrLf & _
                 "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
                 , IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)
        End If

        Screen.MousePointer = vbDefault
        CheckCanMakeHidden = False
        Exit Function

      ElseIf (iCount_NotOwner > 0) Then            ' Cannot change utils
      
        COAMsgBox "This " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " cannot be made hidden as it is used in the" & vbCrLf & _
               "following definition(s), of which you are not the" & vbCrLf & _
               "owner :" & vbCrLf & vbCrLf & sDetails_NotOwner, _
               vbExclamation + vbOKOnly, IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)
        Screen.MousePointer = vbDefault
        CheckCanMakeHidden = False
        Exit Function
      
      End If
    
    '*****************************************************
    Case "P"
    '*****************************************************
      '---------------------------------------------------
      ' Check Cross Tabs For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysCrossTab.Name," & _
        "   AsrSysCrossTab.[CrossTabID] AS [ID]," & _
        "   AsrSysCrossTab.Username," & _
        "   COUNT (ASRSYSCrossTabAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysCrossTab" & _
        " LEFT OUTER JOIN ASRSYSCrossTabAccess ON AsrSysCrossTab.crossTabID = ASRSYSCrossTabAccess.ID" & _
        "   AND ASRSYSCrossTabAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCrossTabAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysCrossTab.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysCrossTab.Name," & _
        "   AsrSysCrossTab.crossTabID," & _
        "   AsrSysCrossTab.Username"
      CheckForPicklistsExpressions utlCrossTab, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCrossTabIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these CrossTabs are contained within a batch job
      If Len(Trim(sCrossTabIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCrossTab, _
          sCrossTabIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Custom Reports For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT ASRSysCustomReportsName.Name," & _
        "   ASRSysCustomReportsName.ID," & _
        "   ASRSysCustomReportsName.Username," & _
        "   COUNT (ASRSYSCustomReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSYSCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSYSCustomReportAccess.ID" & _
        "   AND ASRSYSCustomReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCustomReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysCustomReportsName.Picklist = " & CStr(plngID) & _
        "   OR ASRSysCustomReportsName.Parent1Picklist = " & CStr(plngID) & _
        "   OR ASRSysCustomReportsName.Parent2Picklist = " & CStr(plngID) & _
        " GROUP BY ASRSysCustomReportsName.Name," & _
        "   ASRSysCustomReportsName.ID," & _
        "   ASRSysCustomReportsName.Username"
      
      CheckForPicklistsExpressions utlCustomReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCustomReportsIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Custom Reports are contained within a batch job
      If Len(Trim(sCustomReportsIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCustomReport, _
          sCustomReportsIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Calendar Reports For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT ASRSysCalendarReports.Name," & _
        "   ASRSysCalendarReports.ID," & _
        "   ASRSysCalendarReports.Username," & _
        "   COUNT (ASRSYSCalendarReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysCalendarReports" & _
        " LEFT OUTER JOIN ASRSYSCalendarReportAccess ON ASRSysCalendarReports.ID = ASRSYSCalendarReportAccess.ID" & _
        "   AND ASRSYSCalendarReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSCalendarReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysCalendarReports.Picklist = " & CStr(plngID) & _
        " GROUP BY ASRSysCalendarReports.Name," & _
        "   ASRSysCalendarReports.ID," & _
        "   ASRSysCalendarReports.Username"
                
      CheckForPicklistsExpressions utlCalendarReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCalendarReportsIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Custom Reports are contained within a batch job
      If Len(Trim(sCalendarReportsIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCalendarReport, _
          sCalendarReportsIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Record Profile For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT ASRSysRecordProfileName.Name," & _
        "   ASRSysRecordProfileName.recordProfileID AS [ID]," & _
        "   ASRSysRecordProfileName.Username," & _
        "   COUNT (ASRSYSRecordProfileAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysRecordProfileName" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileAccess.ID" & _
        "   AND ASRSYSRecordProfileAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSRecordProfileAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysRecordProfileName.PicklistID = " & CStr(plngID) & _
        " GROUP BY ASRSysRecordProfileName.Name," & _
        "   ASRSysRecordProfileName.recordProfileID," & _
        "   ASRSysRecordProfileName.Username"
      CheckForPicklistsExpressions utlRecordProfile, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sRecordProfileIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
    
      ' Now check that any of these Record Profiles are contained within a batch job
      If Len(Trim(sRecordProfileIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlRecordProfile, _
          sRecordProfileIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Data Transfer For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysDataTransferName.Name," & _
        "   AsrSysDataTransferName.DataTransferID AS [ID]," & _
        "   AsrSysDataTransferName.Username," & _
        "   COUNT (ASRSYSDataTransferAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysDataTransferName" & _
        " LEFT OUTER JOIN ASRSYSDataTransferAccess ON AsrSysDataTransferName.DataTransferID = ASRSYSDataTransferAccess.ID" & _
        "   AND ASRSYSDataTransferAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSDataTransferAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysDataTransferName.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysDataTransferName.Name," & _
        "   AsrSysDataTransferName.DataTransferID," & _
        "   AsrSysDataTransferName.Username"
      
      
      CheckForPicklistsExpressions utlDataTransfer, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sDataTransferIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these DataTransfers are contained within a batch job
      If Len(Trim(sDataTransferIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlDataTransfer, _
          sDataTransferIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Export For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysExportName.Name," & _
        "   AsrSysExportName.ID," & _
        "   AsrSysExportName.Username," & _
        "   COUNT (ASRSYSExportAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysExportName" & _
        " LEFT OUTER JOIN ASRSYSExportAccess ON AsrSysExportName.ID = ASRSYSExportAccess.ID" & _
        "   AND ASRSYSExportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSExportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysExportName.Picklist = " & CStr(plngID) & _
        "   OR AsrSysExportName.Parent1Picklist = " & CStr(plngID) & _
        "   OR AsrSysExportName.Parent2Picklist = " & CStr(plngID) & _
        " GROUP BY AsrSysExportName.Name," & _
        "   AsrSysExportName.ID," & _
        "   AsrSysExportName.Username"
      
      CheckForPicklistsExpressions utlExport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sExportIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Exports are contained within a batch job
      If Len(Trim(sExportIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlExport, _
          sExportIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Global Add For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysGlobalFunctions.Type = 'A' " & _
        "  AND AsrSysGlobalFunctions.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      
      CheckForPicklistsExpressions UtlGlobalAdd, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalAddIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Global Adds are contained within a batch job
      If Len(Trim(sGlobalAddIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs UtlGlobalAdd, _
          sGlobalAddIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Global Update For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysGlobalFunctions.Type = 'U' " & _
        "  AND AsrSysGlobalFunctions.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      
      CheckForPicklistsExpressions utlGlobalUpdate, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalUpdateIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check that any of these Global Updates are contained within a batch job
      If Len(Trim(sGlobalUpdateIDs)) > 0 Then
        ' JPD20011219 Fault 3303
        CheckCanMakeHiddenInBatchJobs utlGlobalUpdate, _
          sGlobalUpdateIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Global Delete For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID AS [ID]," & _
        "   AsrSysGlobalFunctions.Username," & _
        "   COUNT (ASRSYSGlobalAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysGlobalFunctions" & _
        " LEFT OUTER JOIN ASRSYSGlobalAccess ON AsrSysGlobalFunctions.functionID = ASRSYSGlobalAccess.ID" & _
        "   AND ASRSYSGlobalAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSGlobalAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysGlobalFunctions.Type = 'D' " & _
        "  AND AsrSysGlobalFunctions.PicklistID = " & CStr(plngID) & _
        " GROUP BY AsrSysGlobalFunctions.Name," & _
        "   AsrSysGlobalFunctions.functionID," & _
        "   AsrSysGlobalFunctions.Username"
      
      CheckForPicklistsExpressions utlGlobalDelete, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sGlobalDeleteIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check that any of these Global Deletes are contained within a batch job
      If Len(Trim(sGlobalDeleteIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlGlobalDelete, _
          sGlobalDeleteIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Mail Merge For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID AS [ID]," & _
        "   AsrSysMailMergeName.Username," & _
        "   COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysMailMergeName" & _
        " LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.mailMergeID = ASRSYSMailMergeAccess.ID" & _
        "   AND ASRSYSMailMergeAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysMailMergeName.PicklistID = " & plngID & _
        "   AND AsrSysMailMergeName.isLabel = 0" & _
        " GROUP BY AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID," & _
        "   AsrSysMailMergeName.Username"
      
      CheckForPicklistsExpressions utlMailMerge, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sMailMergeIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check if any of these Merges are contained within a batch job
      If Len(Trim(sMailMergeIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlMailMerge, _
          sMailMergeIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Envelopes & Labels For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID AS [ID]," & _
        "   AsrSysMailMergeName.Username," & _
        "   COUNT (ASRSYSMailMergeAccess.Access) AS [nonHiddenCount]" & _
        " FROM AsrSysMailMergeName" & _
        " LEFT OUTER JOIN ASRSYSMailMergeAccess ON AsrSysMailMergeName.mailMergeID = ASRSYSMailMergeAccess.ID" & _
        "   AND ASRSYSMailMergeAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMailMergeAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE AsrSysMailMergeName.PicklistID = " & plngID & _
        "   AND AsrSysMailMergeName.isLabel = 1" & _
        " GROUP BY AsrSysMailMergeName.Name," & _
        "   AsrSysMailMergeName.MailMergeID," & _
        "   AsrSysMailMergeName.Username"
      
      CheckForPicklistsExpressions utlLabel, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sLabelIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner

      ' Now check if any of these Merges are contained within a batch job
      If Len(Trim(sLabelIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlLabel, _
          sLabelIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If

      '---------------------------------------------------
      ' Check Match Report For This Picklist
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 0 " & _
        "  AND (ASRSysMatchReportName.table1Picklist = " & CStr(plngID) & _
        "  OR ASRSysMatchReportName.table2Picklist = " & CStr(plngID) & ")" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlMatchReport, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sMatchReportIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sMatchReportIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlMatchReport, _
          sMatchReportIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Succession Planning For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 1 " & _
        "  AND (ASRSysMatchReportName.table1Picklist = " & CStr(plngID) & _
        "  OR ASRSysMatchReportName.table2Picklist = " & CStr(plngID) & ")" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlSuccession, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sSuccessionPlanningIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sSuccessionPlanningIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlSuccession, _
          sSuccessionPlanningIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Check Career Progression For This Filter
      '---------------------------------------------------
      sSQL = "SELECT ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID AS [ID]," & _
        "   ASRSysMatchReportName.Username," & _
        "   COUNT (ASRSYSMatchReportAccess.Access) AS [nonHiddenCount]" & _
        " FROM ASRSysMatchReportName" & _
        " LEFT OUTER JOIN ASRSYSMatchReportAccess ON ASRSysMatchReportName.MatchReportID = ASRSYSMatchReportAccess.ID" & _
        "   AND ASRSYSMatchReportAccess.access <> '" & ACCESS_HIDDEN & "'" & _
        "   AND ASRSYSMatchReportAccess.groupName NOT IN (SELECT sysusers.name" & _
        "     FROM sysusers" & _
        "     INNER JOIN ASRSysGroupPermissions ON sysusers.name = ASRSysGroupPermissions.groupName" & _
        "       AND ASRSysGroupPermissions.permitted = 1" & _
        "     INNER JOIN ASRSysPermissionItems ON (ASRSysGroupPermissions.itemID  = ASRSysPermissionItems.itemID" & _
        "       AND (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
        "       OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER'))" & _
        "     INNER JOIN ASRSysPermissionCategories ON (ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
        "       AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & _
        "     WHERE sysusers.uid = sysusers.gid" & _
        "       AND sysusers.uid <> 0)" & _
        " WHERE ASRSysMatchReportName.matchReportType = 2 " & _
        "  AND (ASRSysMatchReportName.table1Picklist = " & CStr(plngID) & _
        "  OR ASRSysMatchReportName.table2Picklist = " & CStr(plngID) & ")" & _
        " GROUP BY ASRSysMatchReportName.Name," & _
        "   ASRSysMatchReportName.MatchReportID," & _
        "   ASRSysMatchReportName.Username"
      CheckForPicklistsExpressions utlCareer, _
        sSQL, _
        pstrUser, _
        iCount_Owner, _
        sDetails_Owner, _
        sCareerProgressionIDs, _
        iCount_NotOwner, _
        sDetails_NotOwner
      
      ' Now check if any of these Match Reports are contained within a batch job
      If Len(Trim(sCareerProgressionIDs)) > 0 Then
        CheckCanMakeHiddenInBatchJobs utlCareer, _
          sCareerProgressionIDs, _
          pstrUser, _
          iCount_Owner, _
          sBatchJobDetails_Owner, _
          sBatchJobIDs, _
          sBatchJobDetails_NotOwner, _
          fBatchJobsOK, _
          sBatchJobDetails_ScheduledForOtherUsers, _
          sBatchJobScheduledUserGroups
      End If
      
      '---------------------------------------------------
      ' Ok, all relevant utility definitions have now been checked, so check
      ' the counts and act accordingly
      '---------------------------------------------------
      If (iCount_Owner = 0) And _
        (iCount_NotOwner = 0) And _
        fBatchJobsOK And _
        (Len(sBatchJobDetails_Owner) = 0) Then
          
        CheckCanMakeHidden = True
        Exit Function
      
      ElseIf (iCount_Owner > 0) And _
        (iCount_NotOwner = 0) And _
        fBatchJobsOK Then
        ' Can change utils and no utils
        ' are contained within batch jobs
        ' that cant be changed
'        If COAMsgBox("Changing the selected picklist to hidden will automatically" & vbCrLf & _
'                  "make the following definition(s), of which you are the" & vbCrLf & _
'                  "owner, hidden also:" & vbCrLf & vbCrLf & _
'                  sDetails_Owner & sBatchJobDetails_Owner & vbCrLf & _
'                  "Do you wish to continue ?", vbQuestion + vbYesNo, IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)) _
'                  = vbNo Then
'          Screen.MousePointer = vbDefault
'          CheckCanMakeHidden = False
'          Exit Function
        'NHRD12112004 Fault 6335
        If COAMsgBox("Changing the selected " & IIf(UCase(pstrType) = "F", "filter", "calculation") & " to hidden will automatically " & _
                  "make the following definition(s)," & vbCrLf & " of which you are the owner, hidden also:" & vbCrLf & vbCrLf & _
                  sDetails_Owner & sBatchJobDetails_Owner & vbCrLf & _
                  "Do you wish to continue ?", vbQuestion + vbYesNo, IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)) _
                  = vbNo Then
          Screen.MousePointer = vbDefault
          CheckCanMakeHidden = False
          Exit Function
        Else
          ' Ok, we are continuing, so lets update all those utils to hidden !
          
          ' Cross Tabs
          If Len(Trim(sCrossTabIDs)) > 0 Then
            HideUtilities utlCrossTab, sCrossTabIDs
            Call UtilUpdateLastSavedMultiple(utlCrossTab, sCrossTabIDs)
          End If

          ' Custom Reports
          If Len(Trim(sCustomReportsIDs)) > 0 Then
            HideUtilities utlCustomReport, sCustomReportsIDs
            Call UtilUpdateLastSavedMultiple(utlCustomReport, sCustomReportsIDs)
          End If

          ' Calendar Reports
          If Len(Trim(sCalendarReportsIDs)) > 0 Then
            HideUtilities utlCalendarReport, sCalendarReportsIDs
            Call UtilUpdateLastSavedMultiple(utlCalendarReport, sCalendarReportsIDs)
          End If

          ' Record Profile
          If Len(Trim(sRecordProfileIDs)) > 0 Then
            HideUtilities utlRecordProfile, sRecordProfileIDs
            Call UtilUpdateLastSavedMultiple(utlRecordProfile, sRecordProfileIDs)
          End If

          ' Data Transfer
          If Len(Trim(sDataTransferIDs)) > 0 Then
            HideUtilities utlDataTransfer, sDataTransferIDs
            Call UtilUpdateLastSavedMultiple(utlDataTransfer, sDataTransferIDs)
          End If

          ' Export
          If Len(Trim(sExportIDs)) > 0 Then
            HideUtilities utlExport, sExportIDs
            Call UtilUpdateLastSavedMultiple(utlExport, sExportIDs)
          End If

          ' Global Add
          If Len(Trim(sGlobalAddIDs)) > 0 Then
            HideUtilities UtlGlobalAdd, sGlobalAddIDs
            Call UtilUpdateLastSavedMultiple(UtlGlobalAdd, sGlobalAddIDs)
          End If

          ' Global Update
          If Len(Trim(sGlobalUpdateIDs)) > 0 Then
            HideUtilities utlGlobalUpdate, sGlobalUpdateIDs
            Call UtilUpdateLastSavedMultiple(utlGlobalUpdate, sGlobalUpdateIDs)
          End If

          ' Global Delete
          If Len(Trim(sGlobalDeleteIDs)) > 0 Then
            HideUtilities utlGlobalDelete, sGlobalDeleteIDs
            Call UtilUpdateLastSavedMultiple(utlGlobalDelete, sGlobalDeleteIDs)
          End If

          ' Mail Merge
          If Len(Trim(sMailMergeIDs)) > 0 Then
            HideUtilities utlMailMerge, sMailMergeIDs
            Call UtilUpdateLastSavedMultiple(utlMailMerge, sMailMergeIDs)
          End If
          
          ' Envelopes & Labels
          If Len(Trim(sLabelIDs)) > 0 Then
            HideUtilities utlLabel, sLabelIDs
            Call UtilUpdateLastSavedMultiple(utlLabel, sLabelIDs)
          End If
          
          ' Match Reports
          If Len(Trim(sMatchReportIDs)) > 0 Then
            HideUtilities utlMatchReport, sMatchReportIDs
            Call UtilUpdateLastSavedMultiple(utlMatchReport, sMatchReportIDs)
          End If
          
          ' Succession Planning
          If Len(Trim(sSuccessionPlanningIDs)) > 0 Then
            HideUtilities utlSuccession, sSuccessionPlanningIDs
            Call UtilUpdateLastSavedMultiple(utlSuccession, sSuccessionPlanningIDs)
          End If
          
          ' Career Progression
          If Len(Trim(sCareerProgressionIDs)) > 0 Then
            HideUtilities utlCareer, sCareerProgressionIDs
            Call UtilUpdateLastSavedMultiple(utlCareer, sCareerProgressionIDs)
          End If
          
          ' Batch Jobs
          If Len(Trim(sBatchJobIDs)) > 0 Then
            HideUtilities utlBatchJob, sBatchJobIDs
            Call UtilUpdateLastSavedMultiple(utlBatchJob, sBatchJobIDs)
          End If
          
          ' Ok, all done, so exit now
          CheckCanMakeHidden = True
          Exit Function
        End If
      
      ElseIf (iCount_Owner > 0) And _
        (iCount_NotOwner = 0) And _
        (Not fBatchJobsOK) Then
        ' Can change utils but abort cos those
        ' utils are in batch jobs which cannot
        ' be changed
        If Len(sBatchJobDetails_ScheduledForOtherUsers) > 0 Then
          COAMsgBox "This picklist cannot be made hidden as it is used in " & vbCrLf & _
                 "definition(s) which are included in the following" & vbCrLf & _
                 "batch jobs which are scheduled to be run by other user groups :" & vbCrLf & vbCrLf & sBatchJobDetails_ScheduledForOtherUsers, vbExclamation + vbOKOnly _
                 , IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)
        Else
          COAMsgBox "This picklist cannot be made hidden as it is used in " & vbCrLf & _
                 "definition(s) which are included in the following" & vbCrLf & _
                 "batch jobs of which you are not the owner :" & vbCrLf & vbCrLf & sBatchJobDetails_NotOwner, vbExclamation + vbOKOnly _
                 , IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)
        End If
        Screen.MousePointer = vbDefault
        CheckCanMakeHidden = False
        Exit Function

      ElseIf (iCount_NotOwner > 0) Then
        ' Cannot change utils
        COAMsgBox "This picklist cannot be made hidden as it is used in the" & vbCrLf & _
               "following definition(s), of which you are not the" & vbCrLf & _
               "owner :" & vbCrLf & vbCrLf & sDetails_NotOwner, _
               vbExclamation + vbOKOnly, IIf(Len(pstrCaption) = 0, "OpenHR - Data Manager", pstrCaption)
        Screen.MousePointer = vbDefault
        CheckCanMakeHidden = False
        Exit Function
      
      End If
    
  End Select
  
End Function


Public Function DateColumn(strType As String, lngTableID As Long, lngColumnID As Long) As Boolean

  'MH20000705
  Dim objCalcExpr As New clsExprExpression
  'Dim rsTemp As ADODB.Recordset
  'Dim sSQL As String

  DateColumn = False


  Select Case strType
  Case "C"      'Column
    DateColumn = (GetDataType(lngTableID, lngColumnID) = sqlDate)
  
  Case Else      'Calculation
    Set objCalcExpr = New clsExprExpression
    objCalcExpr.Initialise lngTableID, lngColumnID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED
    objCalcExpr.ConstructExpression
    
    'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
    ' when all we need is the expression return type (time saving measure).
    objCalcExpr.ValidateExpression True, True

    DateColumn = (objCalcExpr.ReturnType = giEXPRVALUE_DATE)
    Set objCalcExpr = Nothing

  End Select

End Function

Public Function NumericColumn(strType As String, lngTableID As Long, lngColumnID As Long) As Boolean

  ' RH 02/10/00

  NumericColumn = False
  
  Dim objCalcExpr As New clsExprExpression
  
  Select Case strType
  
    Case "C"      'Column
      
      NumericColumn = (GetDataType(lngTableID, lngColumnID) = sqlNumeric) Or _
                      GetDataType(lngTableID, lngColumnID) = sqlInteger
  
    Case Else
      'TM20020417 Fault 3772 - can also check expressions for the Return Type
      Set objCalcExpr = New clsExprExpression
      objCalcExpr.Initialise lngTableID, lngColumnID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED
      objCalcExpr.ConstructExpression
      
      'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
      ' when all we need is the expression return type (time saving measure).
      objCalcExpr.ValidateExpression True, True
  
      NumericColumn = (objCalcExpr.ReturnType = giEXPRVALUE_NUMERIC)
      Set objCalcExpr = Nothing

  End Select
  
End Function

Public Function BitColumn(strType As String, lngTableID As Long, lngColumnID As Long) As Boolean

  'RH20000713
  Dim objCalcExpr As New clsExprExpression
  'Dim rsTemp As ADODB.Recordset
  'Dim sSQL As String

  BitColumn = False

  Select Case strType
  Case "C"      'Column
    BitColumn = (GetDataType(lngTableID, lngColumnID) = sqlBoolean)
  
  Case Else      'Calculation
    Set objCalcExpr = New clsExprExpression
    objCalcExpr.Initialise lngTableID, lngColumnID, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_UNDEFINED
    objCalcExpr.ConstructExpression
    
    'JPD 20031212 Pass optional parameter to avoid creating the expression SQL code
    ' when all we need is the expression return type (time saving measure).
    objCalcExpr.ValidateExpression True, True

    BitColumn = (objCalcExpr.ReturnType = giEXPRVALUE_LOGIC)
    Set objCalcExpr = Nothing

  End Select

End Function

Public Function InsertTableRecord(psInsertSQL As String, _
  plngTableID As Long, _
  plngNewRecordID As Long) As Boolean
  ' Insert the given record in the given table.
  On Error GoTo ErrorTrap
  
  Dim fInsertedOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sSQL As String
  Dim sErrorMsg As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC
  Dim fRecDescExists As Boolean
  Dim sFirstChar As String
  
  Const iRETRIES = 5
  Const iPAUSE = 5000
  
  iOldCursorLocation = gADOCon.CursorLocation
  fInsertedOK = True
  iRetryCount = 0
            
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "sp_ASRInsertNewRecord"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
              
      Set pmADO = .CreateParameter("newID", adInteger, adParamOutput)
      .Parameters.Append pmADO
    
      Set pmADO = .CreateParameter("insertString", adLongVarChar, adParamInput, -1)
      .Parameters.Append pmADO
      pmADO.Value = psInsertSQL
              
      Set pmADO = Nothing
  
      fDeadlock = True
      Do While fDeadlock
        fDeadlock = False
                
        ' Change the cursor location to 'client' as the errors that might be raised
        ' during the update cannot be read for 'server' cursors.
        gADOCon.Errors.Clear
        gADOCon.CursorLocation = adUseClient
              
        On Error GoTo InsertErrorTrap
InsertDeadlockRecoveryPoint:
        cmADO.Execute
        On Error GoTo ErrorTrap
    
        ' Restore the original cursor location to the ADO connection object.
        gADOCon.CursorLocation = iOldCursorLocation
              
        ' Check if the update prodcued any errors.
        If gADOCon.Errors.Count > 0 Then
          sErrorMsg = ""
                  
          For Each ADOErr In gADOCon.Errors
            ' If any 'deadlocks' occur, try to save changes again.
            ' Do this a few times and if errors still occur then display a more friendly
            ' error message than the ' deadlock victim' one generated by ODBC.
            If (ADOErr.Number = DEADLOCK_ERRORNUMBER) And _
              (((UCase(Left(ADOErr.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                (UCase(Right(ADOErr.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
              ((UCase(Left(ADOErr.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                (InStr(UCase(ADOErr.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    
              ' The error is for a deadlock.
              ' Sorry about having to use the err.description to trap the error but the err.number
              ' is not specific and MSDN suggests using the err.description.
              If (iRetryCount < iRETRIES) And (gADOCon.Errors.Count = 1) Then
                iRetryCount = iRetryCount + 1
                fDeadlock = True
                ' Pause before resubmitting the SQL command.
                Sleep iPAUSE
              Else
                sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                  "Another user is deadlocking the database. Try saving again."
                fInsertedOK = False
              End If
            
            Else
              sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                ADOErr.Description
              fInsertedOK = False
            End If
          Next ADOErr
    
          gADOCon.Errors.Clear

          If Not fInsertedOK Then
            COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
              sErrorMsg, vbOKOnly + vbExclamation, app.ProductName
          End If
        End If
        
        If fInsertedOK And (Not fDeadlock) Then
          plngNewRecordID = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
        End If
      Loop
    End With
    Set cmADO = Nothing

TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    gADOCon.CursorLocation = iOldCursorLocation
  Else
    gADOCon.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  InsertTableRecord = fInsertedOK
  Exit Function

ErrorTrap:
  fInsertedOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
  Resume TidyUpAndExit

InsertErrorTrap:
  ' If any 'deadlocks' occur, try to save changes again.
  ' Do this a few times and if errors still occur then display a more friendly
  ' error message than the 'deadlock victim' one generated by ODBC.
  If (Err.Number = DEADLOCK_ERRORNUMBER) And _
    (((UCase(Left(Err.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
      (UCase(Right(Err.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
    ((UCase(Left(Err.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
      (InStr(UCase(Err.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    ' The error is for a deadlock.
    ' Sorry about having to use the err.description to trap the error but the err.number
    ' is not specific and MSDN suggests using the err.description.
    If iRetryCount < iRETRIES Then
      iRetryCount = iRetryCount + 1
      ' Pause before resubmitting the SQL command.
      Sleep iPAUSE
      Resume InsertDeadlockRecoveryPoint
    Else
      fInsertedOK = False
      gobjProgress.CloseProgress
      COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        "Another user is deadlocking the database. Try saving again.", _
        vbExclamation + vbOKOnly, Application.Name
      Resume TidyUpAndExit
    End If
  Else
    fInsertedOK = False
    'JPD 20030904 Fault 5479
    gobjProgress.CloseProgress
    
    'JPD 20110705 HRPRO-1572
    ' Now get validation failure message prefixed woth <record description> and <line of hyphens>.
    ' Only add extra carriage return if required (ie. if there is a record description).
    fRecDescExists = False
    sFirstChar = Left(Err.Description, 1)
    If Len(sFirstChar) > 0 Then
      fRecDescExists = (Asc(sFirstChar) <> 13)
    End If
    
    COAMsgBox "The new record could not be created." & _
      IIf(fRecDescExists, vbCrLf & vbCrLf, "") & _
      ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
    Resume TidyUpAndExit
  End If
  
End Function
Public Function DeleteTableRecord(plngTableID As Long, _
  psRealSource As String, _
  plngRecordID As Long) As Boolean
  ' Insert the given record in the given table.
  On Error GoTo ErrorTrap
  
  Dim fDeletedOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sSQL As String
  Dim sErrorMsg As String
  'Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC

  Const iRETRIES = 5
  Const iPAUSE = 5000
  
  iOldCursorLocation = gADOCon.CursorLocation
  fDeletedOK = True
  iRetryCount = 0
  
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "sp_ASRDeleteRecord"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
              
      Set pmADO = .CreateParameter("resultCode", adInteger, adParamOutput)
      .Parameters.Append pmADO
      
      Set pmADO = .CreateParameter("tableID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngTableID
      
      Set pmADO = .CreateParameter("realSource", adVarChar, adParamInput, 255)
      .Parameters.Append pmADO
      pmADO.Value = psRealSource
      
      Set pmADO = .CreateParameter("id", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngRecordID
              
      Set pmADO = Nothing
  
      fDeadlock = True
      Do While fDeadlock
        fDeadlock = False
                
        ' Change the cursor location to 'client' as the errors that might be raised
        ' during the update cannot be read for 'server' cursors.
        gADOCon.Errors.Clear
        gADOCon.CursorLocation = adUseClient
              
        On Error GoTo DeleteErrorTrap
DeleteDeadlockRecoveryPoint:
        cmADO.Execute
        On Error GoTo ErrorTrap
    
        ' Restore the original cursor location to the ADO connection object.
        gADOCon.CursorLocation = iOldCursorLocation
              
        ' Check if the update prodcued any errors.
        If gADOCon.Errors.Count > 0 Then
          sErrorMsg = ""
                  
          For Each ADOErr In gADOCon.Errors
            ' If any 'deadlocks' occur, try to save changes again.
            ' Do this a few times and if errors still occur then display a more friendly
            ' error message than the ' deadlock victim' one generated by ODBC.
            If (ADOErr.Number = DEADLOCK_ERRORNUMBER) And _
              (((UCase(Left(ADOErr.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                (UCase(Right(ADOErr.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
              ((UCase(Left(ADOErr.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                (InStr(UCase(ADOErr.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    
              ' The error is for a deadlock.
              ' Sorry about having to use the err.description to trap the error but the err.number
              ' is not specific and MSDN suggests using the err.description.
              If (iRetryCount < iRETRIES) And (gADOCon.Errors.Count = 1) Then
                iRetryCount = iRetryCount + 1
                fDeadlock = True
                ' Pause before resubmitting the SQL command.
                Sleep iPAUSE
              Else
                sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                  "Another user is deadlocking the database. Try saving again."
                fDeletedOK = False
              End If
            Else
              sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                ADOErr.Description
              fDeletedOK = False
            End If
          Next ADOErr
    
          gADOCon.Errors.Clear

          If Not fDeletedOK Then
            COAMsgBox "The record could not be deleted." & vbCrLf & vbCrLf & _
              sErrorMsg, vbOKOnly + vbExclamation, app.ProductName
          End If
        Else
          Select Case .Parameters(0).Value
            Case 2 ' Record has been amended AND is no longer in the given table/view.
              COAMsgBox "The record has been amended by another user and is no longer in the current view.", vbExclamation, app.ProductName
              
            Case Else ' Record deleted OK.
          End Select
        End If
      Loop
    End With
    Set cmADO = Nothing
    
TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    gADOCon.CursorLocation = iOldCursorLocation
  Else
    gADOCon.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  DeleteTableRecord = fDeletedOK
  Exit Function

ErrorTrap:
  fDeletedOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
  Resume TidyUpAndExit

DeleteErrorTrap:
  ' If any 'deadlocks' occur, try to save changes again.
  ' Do this a few times and if errors still occur then display a more friendly
  ' error message than the 'deadlock victim' one generated by ODBC.
  If (Err.Number = DEADLOCK_ERRORNUMBER) And _
    (((UCase(Left(Err.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
      (UCase(Right(Err.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
    ((UCase(Left(Err.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
      (InStr(UCase(Err.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    ' The error is for a deadlock.
    ' Sorry about having to use the err.description to trap the error but the err.number
    ' is not specific and MSDN suggests using the err.description.
    If iRetryCount < iRETRIES Then
      iRetryCount = iRetryCount + 1
      ' Pause before resubmitting the SQL command.
      Sleep iPAUSE
      Resume DeleteDeadlockRecoveryPoint
    Else
      fDeletedOK = False
      COAMsgBox "The record could not be deleted." & vbCrLf & vbCrLf & _
        "Another user is deadlocking the database. Try saving again.", _
        vbExclamation + vbOKOnly, Application.Name
      Resume TidyUpAndExit
    End If
  Else
    fDeletedOK = False
    COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
    Resume TidyUpAndExit
  End If
  
End Function

Public Function UpdateTableRecord(psUpdateSQL As String, _
  plngTableID As Long, _
  psRealSource As String, _
  plngRecordID As Long, _
  Optional pvTimestamp As Variant) As Boolean
  ' Update the given table with the given SQL update command.
  On Error GoTo ErrorTrap
  
  Dim fUpdatedOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sSQL As String
  Dim sErrorMsg As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC
  Dim fRecDescExists As Boolean
  Dim sFirstChar As String

  Const iRETRIES = 5
  Const iPAUSE = 5000
  
  iOldCursorLocation = gADOCon.CursorLocation
  fUpdatedOK = True
  iRetryCount = 0
  
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "sp_ASRUpdateRecord"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
              
      Set pmADO = .CreateParameter("resultCode", adInteger, adParamOutput)
      .Parameters.Append pmADO
            
      Set pmADO = .CreateParameter("updateString", adLongVarChar, adParamInput, VARCHAR_MAX_Size)
      .Parameters.Append pmADO
      pmADO.Value = psUpdateSQL
              
      Set pmADO = .CreateParameter("tableID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngTableID
            
      Set pmADO = .CreateParameter("realSource", adVarChar, adParamInput, 255)
      .Parameters.Append pmADO
      pmADO.Value = psRealSource
              
      Set pmADO = .CreateParameter("id", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngRecordID
              
      Set pmADO = .CreateParameter("timestamp", adInteger, adParamInput)
      .Parameters.Append pmADO
      If IsMissing(pvTimestamp) Then
        pmADO.Value = Null
      Else
        pmADO.Value = CLng(pvTimestamp)
      End If

      Set pmADO = Nothing
  
      fDeadlock = True
      Do While fDeadlock
        fDeadlock = False
                
        ' Change the cursor location to 'client' as the errors that might be raised
        ' during the update cannot be read for 'server' cursors.
        gADOCon.Errors.Clear
        gADOCon.CursorLocation = adUseClient
              
        On Error GoTo UpdateErrorTrap
UpdateDeadlockRecoveryPoint:
        cmADO.Execute
        On Error GoTo ErrorTrap
    
        ' Restore the original cursor location to the ADO connection object.
        gADOCon.CursorLocation = iOldCursorLocation
              
        ' Check if the update prodcued any errors.
        If gADOCon.Errors.Count > 0 Then
          sErrorMsg = ""
                  
          For Each ADOErr In gADOCon.Errors
            ' If any 'deadlocks' occur, try to save changes again.
            ' Do this a few times and if errors still occur then display a more friendly
            ' error message than the ' deadlock victim' one generated by ODBC.
            If (ADOErr.Number = DEADLOCK_ERRORNUMBER) And _
              (((UCase(Left(ADOErr.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                (UCase(Right(ADOErr.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
              ((UCase(Left(ADOErr.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                (InStr(UCase(ADOErr.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    
              ' The error is for a deadlock.
              ' Sorry about having to use the err.description to trap the error but the err.number
              ' is not specific and MSDN suggests using the err.description.
              If (iRetryCount < iRETRIES) And (gADOCon.Errors.Count = 1) Then
                iRetryCount = iRetryCount + 1
                fDeadlock = True
                ' Pause before resubmitting the SQL command.
                Sleep iPAUSE
              Else
                sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                  "Another user is deadlocking the database. Try saving again."
                fUpdatedOK = False
              End If
            
            Else
              sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
                ADOErr.Description
              fUpdatedOK = False
            End If
          Next ADOErr
    
          gADOCon.Errors.Clear

          If Not fUpdatedOK Then
            COAMsgBox "The record could not be updated." & vbCrLf & vbCrLf & _
              sErrorMsg, vbOKOnly + vbExclamation, app.ProductName
          End If
        End If
        
        If fUpdatedOK And (Not fDeadlock) Then
          Select Case .Parameters(0).Value
            Case 1 ' Record changed by another user, and still in the current table/view.
              COAMsgBox "The record has been amended by another user and will be refreshed.", vbExclamation, app.ProductName
            Case 2 ' Record changed by another user, and is no longer in the current table/view.
              COAMsgBox "The record has been amended by another user and is no longer in the current view.", vbExclamation, app.ProductName
            Case 3 ' Record deleted by another user.
              COAMsgBox "The record has been deleted by another user.", vbExclamation, app.ProductName
            Case Else ' Record updated OK.
              
          End Select
        End If
      Loop
    End With
    Set cmADO = Nothing

TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    gADOCon.CursorLocation = iOldCursorLocation
  Else
    gADOCon.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  UpdateTableRecord = fUpdatedOK
  Exit Function

ErrorTrap:
  fUpdatedOK = False
  COAMsgBox ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
  Resume TidyUpAndExit

UpdateErrorTrap:
  ' If any 'deadlocks' occur, try to save changes again.
  ' Do this a few times and if errors still occur then display a more friendly
  ' error message than the 'deadlock victim' one generated by ODBC.
  If (Err.Number = DEADLOCK_ERRORNUMBER) And _
    (((UCase(Left(Err.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
      (UCase(Right(Err.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
    ((UCase(Left(Err.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
      (InStr(UCase(Err.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    ' The error is for a deadlock.
    ' Sorry about having to use the err.description to trap the error but the err.number
    ' is not specific and MSDN suggests using the err.description.
    If iRetryCount < iRETRIES Then
      iRetryCount = iRetryCount + 1
      ' Pause before resubmitting the SQL command.
      Sleep iPAUSE
      Resume UpdateDeadlockRecoveryPoint
    Else
      fUpdatedOK = False
      gobjProgress.CloseProgress
      COAMsgBox "The record could not be updated." & vbCrLf & vbCrLf & _
        "Another user is deadlocking the database. Try saving again.", _
        vbExclamation + vbOKOnly, Application.Name
      Resume TidyUpAndExit
    End If
  Else
    fUpdatedOK = False
    'JPD 20030904 Fault 5479
    gobjProgress.CloseProgress
    
    'JPD 20110705 HRPRO-1572
    ' Now get validation failure message prefixed woth <record description> and <line of hyphens>.
    ' Only add extra carriage return if required (ie. if there is a record description).
    fRecDescExists = False
    sFirstChar = Left(Err.Description, 1)
    If Len(sFirstChar) > 0 Then
      fRecDescExists = (Asc(sFirstChar) <> 13)
    End If

    COAMsgBox "The record could not be updated." & _
      IIf(fRecDescExists, vbCrLf & vbCrLf, "") & _
      ODBC.FormatError(Err.Description), vbExclamation + vbOKOnly, Application.Name
    Resume TidyUpAndExit
  End If
  
End Function

Public Function ExecuteSql(psSQL As String, psErrorMsg As String) As Boolean
  ' Update the given table with the given SQL update command.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fDeadlock As Boolean
  Dim iRetryCount As Integer
  Dim iOldCursorLocation As Integer
  Dim sErrorMsg As String
  Dim ADOErr As ADODB.Error
  Dim ODBC As New ODBC

  Const iRETRIES = 5
  Const iPAUSE = 5000
  
  iOldCursorLocation = gADOCon.CursorLocation
  fOK = True
  iRetryCount = 0
  
  fDeadlock = True
  Do While fDeadlock
    fDeadlock = False
    
    ' Change the cursor location to 'client' as the errors that might be raised
    ' during the update cannot be read for 'server' cursors.
    gADOCon.Errors.Clear
    gADOCon.CursorLocation = adUseClient
              
    On Error GoTo ExecuteErrorTrap
DeadlockRecoveryPoint:
    gADOCon.Execute psSQL, , adCmdText + adExecuteNoRecords
    On Error GoTo ErrorTrap
    
    ' Restore the original cursor location to the ADO connection object.
    gADOCon.CursorLocation = iOldCursorLocation
              
    ' Check if the execution produced any errors.
    If gADOCon.Errors.Count > 0 Then
      sErrorMsg = ""
                  
      For Each ADOErr In gADOCon.Errors
        ' If any 'deadlocks' occur, try to run the execution again.
        ' Do this a few times and if errors still occur then display a more friendly
        ' error message than the ' deadlock victim' one generated by ODBC.
        If (ADOErr.Number = DEADLOCK_ERRORNUMBER) And _
          (((UCase(Left(ADOErr.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
            (UCase(Right(ADOErr.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
          ((UCase(Left(ADOErr.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
            (InStr(UCase(ADOErr.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    
          ' The error is for a deadlock.
          ' Sorry about having to use the err.description to trap the error but the err.number
          ' is not specific and MSDN suggests using the err.description.
          If (iRetryCount < iRETRIES) And (gADOCon.Errors.Count = 1) Then
            iRetryCount = iRetryCount + 1
            fDeadlock = True
            ' Pause before resubmitting the SQL command.
            Sleep iPAUSE
          Else
            sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
              "Another user is deadlocking the database."
            fOK = False
          End If
        Else
          sErrorMsg = sErrorMsg & IIf(Len(sErrorMsg) > 0, vbCrLf, "") & _
            ADOErr.Description
          fOK = False
        End If
      Next ADOErr
    
      gADOCon.Errors.Clear
    End If
  Loop
  
TidyUpAndExit:
  If (iOldCursorLocation = adUseClient) Or _
    (iOldCursorLocation = adUseServer) Then
    gADOCon.CursorLocation = iOldCursorLocation
  Else
    gADOCon.CursorLocation = adUseServer
  End If
  Set ODBC = Nothing
  
  If Not fOK Then
    psErrorMsg = sErrorMsg
  End If
  
  ExecuteSql = fOK
  Exit Function

ErrorTrap:
  fOK = False
  sErrorMsg = ODBC.FormatError(Err.Description)
  Resume TidyUpAndExit

ExecuteErrorTrap:
  ' If any 'deadlocks' occur, try to save changes again.
  ' Do this a few times and if errors still occur then display a more friendly
  ' error message than the 'deadlock victim' one generated by ODBC.
  If (Err.Number = DEADLOCK_ERRORNUMBER) And _
    (((UCase(Left(Err.Description, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
      (UCase(Right(Err.Description, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
    ((UCase(Left(Err.Description, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
      (InStr(UCase(Err.Description), DEADLOCK2_MESSAGEEND) > 0))) Then
    ' The error is for a deadlock.
    ' Sorry about having to use the err.description to trap the error but the err.number
    ' is not specific and MSDN suggests using the err.description.
    If iRetryCount < iRETRIES Then
      iRetryCount = iRetryCount + 1
      ' Pause before resubmitting the SQL command.
      Sleep iPAUSE
      Resume DeadlockRecoveryPoint
    Else
      fOK = False
      sErrorMsg = "Another user is deadlocking the database."
      Resume TidyUpAndExit
    End If
  Else
    fOK = False
    sErrorMsg = ODBC.FormatError(Err.Description)
    Resume TidyUpAndExit
  End If
  
End Function


Public Function ConvertNumberForSQL(ByVal strInput As String) As String
  'Get a number in the correct format for a SQL string
  '(e.g. on french systems replace decimal comma for a decimal point)
  'ConvertNumberForSQL = Replace(strInput, UI.GetSystemDecimalSeparator, ".")
  ConvertNumberForSQL = Trim(Replace(strInput, UI.GetSystemThousandSeparator, ""))
  ConvertNumberForSQL = Replace(ConvertNumberForSQL, UI.GetSystemDecimalSeparator, ".")
  
End Function

Public Function ConvertNumberForDisplay(ByVal strInput As String) As String
  'Get a number in the correct format for display
  '(e.g. on french systems replace decimal point for a decimal comma)
  ConvertNumberForDisplay = Replace(strInput, ".", UI.GetSystemDecimalSeparator)
End Function

Public Sub FormatTDBNumberControl(objInput As Object)
  
  'MH20010130
  If TypeOf objInput Is TDBNumber6Ctl.TDBNumber Then
    objInput.Separator = "x"
    objInput.DecimalPoint = UI.GetSystemDecimalSeparator
    objInput.Separator = UI.GetSystemThousandSeparator
  
  ElseIf (TypeOf objInput Is TDBNumberCtrl.TDBNumber) Then
    'Need to set to dummy incase new thousand is same as old decimal (cases error)
    objInput.SepThousand = "x"
    objInput.SepDecimal = UI.GetSystemDecimalSeparator
    objInput.SepThousand = UI.GetSystemThousandSeparator
  
  End If

End Sub


Public Function CalcIsReadOnly(lExprID As Long) As Boolean

  On Error GoTo ErrTrap
  
  Dim rsTemp As ADODB.Recordset
  
  Set rsTemp = GetReadOnlyRecords("SELECT * FROM ASRSysExpressions WHERE ExprID = " & CStr(lExprID))
  
  
  'MH20010201 fault 1776
  'If rsTemp.Fields("Access").Value = "RO" Then
  '  CalcIsReadOnly = True
  'End If
  CalcIsReadOnly = (rsTemp.Fields("Access").Value <> ACCESS_READWRITE And _
                    UCase(rsTemp.Fields("Username").Value) <> UCase(gsUserName))

ExitPoint:

  Set rsTemp = Nothing
  Exit Function
  
ErrTrap:
  
  COAMsgBox "Error determining Access rights of selected calc:" & vbCrLf & Err.Description, vbExclamation + vbOKOnly, app.Title
  Resume ExitPoint
End Function

Public Function IsColumnAudited(ByVal plngColumnID As Long) As Boolean

  ' Returns whether the passed in column is an audited field or not

    Dim sSQL As String
    Dim rsTemp As ADODB.Recordset
    
    sSQL = "Select audit From ASRSysColumns Where columnid = " & plngColumnID
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
    If Not rsTemp.BOF And Not rsTemp.EOF Then
      IsColumnAudited = rsTemp(0)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

End Function

Public Function IsColumnReadOnly(ByVal plngColumnID As Long) As Boolean

  ' Returns whether the passed in column is an audited field or not

    Dim sSQL As String
    Dim rsTemp As ADODB.Recordset
    
    IsColumnReadOnly = False
    sSQL = "Select readonly From ASRSysColumns Where columnid = " & plngColumnID
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
    If Not rsTemp.BOF And Not rsTemp.EOF Then
      IsColumnReadOnly = rsTemp(0).Value
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

End Function

Public Function ColumnLastChangeDate(ByVal plngRecordID, ByVal psTableName, ByVal psColumnName As String, ByVal pbUpdateLog As Boolean) As Date

  ' plngColumnID = The column ID we want to check on
  ' pbUpdateLog = Do we flag that this audit record has already been exported
    
    Dim sSQL As String
    Dim rsTemp As ADODB.Recordset
    
    sSQL = "Select Top 1 DateTimeStamp, ID From ASRSysAuditTrail" _
        & " Where columnname = '" & psColumnName & "' And tablename ='" & psTableName _
        & "' and recordid = " & plngRecordID _
        & "AND CMGCommitDate Is Null Order By 1 Desc"
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
    If Not rsTemp.BOF And Not rsTemp.EOF Then
      ColumnLastChangeDate = rsTemp(0)
    Else
      ColumnLastChangeDate = #12/31/9999#
    End If
    
    ' Update the audit log to say these have been exported
    If pbUpdateLog = True Then
      sSQL = "Update ASRSysAuditTrail Set CMGExportDate = getdate()" _
        & " Where columnname = '" & psColumnName & "' And tablename ='" & psTableName & "' and recordid = " & plngRecordID _
        & " And CMGExportDate Is Null"
      datData.ExecuteSql sSQL
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing

End Function

'Public Function IsCMGEnabled() As Boolean
'
'  ' Return TRUE if the Personnel module has been enabled.
'  Dim lngCustNo As Long
'  Dim sSQL As String
'  Dim sAuth As String
'  Dim objLicence As ASRLicense.CLicense
'  Dim rsConfig As ADODB.Recordset
'
'  ' Get the Customer number and Module authorisation code.
'  sSQL = "SELECT * FROM ASRSysConfig"
'  Set rsConfig = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  If Not rsConfig.BOF And Not rsConfig.EOF Then
'    lngCustNo = IIf(IsNull(rsConfig!CustNo), 0, rsConfig!CustNo)
'    sAuth = IIf(IsNull(rsConfig!ModuleCode), "", rsConfig!ModuleCode)
'  End If
'  rsConfig.Close
'  Set rsConfig = Nothing
'
'  ' Validate the Personnel module.
'  Set objLicence = New ASRLicense.CLicense
'  'IsCMGEnabled = objLicence.GetModule(CMG, sAuth, lngCustNo)
'  IsCMGEnabled = True
'  Set objLicence = Nothing
'
'End Function

Public Function CMGCommit() As Boolean

  ' Commit the changes to the audit log
  Dim sSQL As String
  Dim bOK As Boolean
 
  On Error GoTo ErrTrap
  bOK = True
  
  ' Update the audit log to say these have been exported
  sSQL = "Update asrsysAuditTrail Set CMGCommitDate = Convert(smalldatetime,getdate()) " _
    & "Where CMGCommitDate Is Null and CMGExportDate Is Not Null"
  datData.ExecuteSql sSQL
  
ExitPoint:
  CMGCommit = bOK
  Exit Function
  
ErrTrap:
  bOK = False
  Resume ExitPoint

End Function

Public Function GetRecordOrderNumber(pvSearchValue As Variant, _
  pstrOrderName As String, _
  psTableName As String, _
  mbFirstColumnAscending As Boolean, _
  piDataType As ADODB.DataTypeEnum) As Long

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.GetRecordOrderNumber(pvSearchValue,pstrOrderName,psTableName)", Array(pvSearchValue, pstrOrderName, psTableName)

  ' Return the record number from the given order.
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
  Dim lngCount As Long
    
  lngCount = 0
    
  ' Blank
  If IsNull(pvSearchValue) Or IsEmpty(pvSearchValue) Then
    pvSearchValue = ""
  End If
       
  ' Format Date
  If piDataType = adDBTimeStamp Then
    pvSearchValue = "CONVERT(datetime, '" & Replace(Format(pvSearchValue, "MM/dd/yyyy"), UI.GetSystemDateSeparator, "/") + "')"
  Else
    pvSearchValue = "'" & Replace(pvSearchValue, "'", "''") & "'"
  End If
    
  sSQL = "SELECT COUNT(" & pstrOrderName & ") AS RecordCount FROM " & psTableName _
    & " WHERE " & pvSearchValue & IIf(mbFirstColumnAscending, " > ", " < ") & pstrOrderName _
    & " ORDER BY 1"
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  If Not (rsTemp.EOF And rsTemp.BOF) Then
    rsTemp.MoveFirst
    lngCount = rsTemp.Fields("RecordCount")
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing

  GetRecordOrderNumber = lngCount

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

Public Function GetManualSummaryBreak(ByVal plngTableID As Long) As Boolean
  Dim rsTable As ADODB.Recordset
  Dim sSQL As String
    
  sSQL = "SELECT ManualSummaryColumnBreaks " & _
    " FROM ASRSysTables " & _
    " WHERE tableID=" & plngTableID
        
  Set rsTable = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsTable
    If Not (.BOF And .EOF) Then
      GetManualSummaryBreak = IIf(IsNull(.Fields(0).Value), False, .Fields(0).Value)
    Else
      GetManualSummaryBreak = False
    End If
    
    .Close
  End With
    
  Set rsTable = Nothing
  
End Function

Public Function GetColumnTableName(plngColumnID As Long) As String
  
  ' Return the table id of the given column.
  Dim sSQL As String
  Dim rsData As ADODB.Recordset

  sSQL = "SELECT tableName" & _
    " FROM ASRSysColumns " & _
    " JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID" & _
    " WHERE columnID = " & Trim(Str(plngColumnID))
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
    GetColumnTableName = rsData!TableName
  Else
    GetColumnTableName = ""
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Public Function GetColumnTrimmingType(plngColumnID As Long) As TrimmingTypes

  ' Return the table id of the given column.
  Dim sSQL As String
  Dim rsData As ADODB.Recordset

  sSQL = "SELECT Trimming FROM ASRSysColumns" _
    & " WHERE columnID = " & Trim(Str(plngColumnID))
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
    GetColumnTrimmingType = IIf(Not IsNull(rsData!Trimming), rsData!Trimming, giTRIMMING_NONE)
  Else
    GetColumnTrimmingType = giTRIMMING_NONE
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Public Function EnableUDFFunctions() As Boolean

  Dim sSQL As String
  Dim rsUser As ADODB.Recordset
  
  sSQL = "exec master..xp_msver"
  
  Set rsUser = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  rsUser.MoveNext
  
  Select Case Val(rsUser(3))
    Case Is >= 8
      EnableUDFFunctions = True
    Case Else
      EnableUDFFunctions = False
  End Select
  
  rsUser.Close
  Set rsUser = Nothing
    
End Function


Public Function FilterUDFs(plngFilterID As Long, pastrUDFs() As String) As Boolean

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsGeneral.FilterUDFs(plngFilterID,pastrUDFs)", Array(plngFilterID, pastrUDFs())
  
  ' Return a string describing the record IDs from the given table
  ' that satisfy the given criteria.
  Dim fOK As Boolean
  Dim objExpr As clsExprExpression

  fOK = True

  Set objExpr = New clsExprExpression
  With objExpr
    .Initialise 0, plngFilterID, giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC
    .UDFFilterCode pastrUDFs(), True, True
  End With
  Set objExpr = Nothing
  
  FilterUDFs = fOK

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function

Public Function DoesColumnUseSeparators(plngColumnID As Long) As Boolean

  ' Returns whether the column uses 1000 separators...
  Dim sSQL As String
  Dim rsData As ADODB.Recordset

  sSQL = "SELECT Use1000Separator FROM ASRSysColumns" _
    & " WHERE columnID = " & Trim(Str(plngColumnID))
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
    DoesColumnUseSeparators = IIf(Not IsNull(rsData!Use1000Separator), rsData!Use1000Separator, False)
  Else
    DoesColumnUseSeparators = False
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Public Function HowManyRowsInALabel(plngLabelID As Long) As Integer

  ' Returns how many rows a label type can show
  Dim sSQL As String
  Dim rsData As ADODB.Recordset

  sSQL = "SELECT LabelHeight, PageHeight, IsEnvelope, StandardFontSize" _
    & " FROM ASRSysLabelTypes" _
    & " WHERE LabelTypeID = " & Trim(Str(plngLabelID))
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
    
    If rsData!IsEnvelope Then
      HowManyRowsInALabel = IIf(Not IsNull(rsData!PageHeight), ((rsData!PageHeight * 28.34375) / (rsData!StandardFontSize + 2)), 0)
    Else
      HowManyRowsInALabel = IIf(Not IsNull(rsData!LabelHeight), ((rsData!LabelHeight * 28.34375) / (rsData!StandardFontSize + 2)), 0)
    End If
  Else
    HowManyRowsInALabel = 0
  End If

  rsData.Close
  Set rsData = Nothing

End Function

' Returns the amount of decimals that are specificed for a column
Public Function GetDecimalsSize(plngColumnID As Long) As Integer

  Dim sSQL As String
  Dim rsData As ADODB.Recordset

  sSQL = "SELECT Decimals FROM ASRSysColumns" _
    & " WHERE columnID = " & Trim(Str(plngColumnID))
  Set rsData = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsData.BOF And Not rsData.EOF Then
    GetDecimalsSize = IIf(Not IsNull(rsData!Decimals), rsData!Decimals, 0)
  Else
    GetDecimalsSize = 0
  End If

  rsData.Close
  Set rsData = Nothing

End Function

Public Function QAddressEnabled() As DataMgr.QAddressTypes

  On Error GoTo QAddressNotInstalled

  Dim bLicensed As Boolean
  Dim strSeedValue As String * 100
  Dim lSearchReturn As Long
  Dim lOpen As Long
  Dim iQAddressType As DataMgr.QAddressTypes
  Dim lngHandle As Long

  bLicensed = IsModuleEnabled(modQAddress)

  ' Calculate the version of Quick Address
  If bLicensed Then

    ' Default Quick Address for this machine
    iQAddressType = QADDRESS_DISABLED

    If GetSystemSetting("Development", "QAddress_Evaluation_Enable", False) Then
      strSeedValue = GetSystemSetting("Development", "QAddress_Evaluation_Seed", "AL1 5ST")
    Else
      strSeedValue = "AL1 5ST"
    End If

    ' Look for Quick Address World Wide
    'iQAddressType = QADDRESS_WORLDWIDE

    ' Look for Quick Address Pro 4
    'iQAddressType = QADDRESS_PRO4

    ' Look for Quick Address Pro 3
    Call QAPro_Close
    Call QAInitialise(1)
    lOpen = QAPro_Open("", "")
    If lOpen = 0 Then
      lSearchReturn = QAPro_Search(MakeQAddressString(strSeedValue))
      If lSearchReturn >= 0 Then
        iQAddressType = QADDRESS_PRO3
        GoTo ExitFunction
      End If
    End If

    ' Look for Quick Address Rapid
    Call QARapid_Close
    lOpen = QARapid_Open("", "")
    If lOpen = 0 Then
      lSearchReturn = QARapid_Search(MakeQAddressString(strSeedValue))
      If lSearchReturn >= 0 Then
        iQAddressType = QADDRESS_RAPID
        GoTo ExitFunction
      End If
    End If
    
    ' NPG20081212 Fault 13364
    ' Look for Quick Address Pro 5+
    lngHandle = 0
    Call QA_Close(lngHandle)
    lOpen = QA_Open("", "", lngHandle)
    If lOpen = 0 Then
      lSearchReturn = QA_Search(lngHandle, strSeedValue)
      If lSearchReturn >= 0 Then
        iQAddressType = QADDRESS_PRO5
        Call QA_Close(lngHandle)
        Call QA_Shutdown
        GoTo ExitFunction
      End If
    End If
  Else
    iQAddressType = QADDRESS_DISABLED
  End If
  
  Call QA_Close(lngHandle)
  Call QA_Shutdown

  
  
  ' JDM - 25/02/3004 - Fault 8000 - Not warning if Quick Address is not installed
  If iQAddressType = QADDRESS_DISABLED And bLicensed Then
    If GetPCSetting("QAddress", "ShowError" & CStr(app.Major & "." & app.Minor), True) = True Then
      If Not ASRDEVELOPMENT And Not gblnBatchJobsOnly Then
        COAMsgBox "OpenHR is configured to use Quick Address Postcode recognition," & vbCrLf & _
               "but one or more of the required files are missing from your PC." & vbCrLf & vbCrLf & _
               "OpenHR will continue, but all Quick Address functionality will be disabled." & vbCrLf & vbCrLf & _
               "(This message will not appear again until a new version of OpenHR is installed.)" _
               , vbExclamation + vbOKOnly, Application.Name
        SavePCSetting "QAddress", "ShowError" & CStr(app.Major & "." & app.Minor), False
      End If
    End If
  End If

ExitFunction:
  QAddressEnabled = iQAddressType
  Exit Function

QAddressNotInstalled:
  lSearchReturn = -1
  Resume Next

End Function


' Gets the maximum size of an OLE column, if no maximum then return -1
Public Function GetOLEMaxSize(lColumnID As Long) As Long

  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset

  sSQL = "Select MaxOLESize From ASRSysColumns Where columnid = " & lColumnID _
      & " AND OLEType = 2 AND MaxOLESizeEnabled = 1"
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  If Not rsTemp.BOF And Not rsTemp.EOF Then
    GetOLEMaxSize = rsTemp(0)
  Else
    GetOLEMaxSize = -1
  End If
  
  rsTemp.Close
  Set rsTemp = Nothing

End Function

Public Function GetSqlProcessID() As Long

  Dim sSQL As String
  Dim rsUser As ADODB.Recordset
  
  sSQL = "SELECT @@SPID"
  
  Set rsUser = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  If Not (rsUser.EOF And rsUser.BOF) Then
    GetSqlProcessID = rsUser.Fields(0).Value
  Else
    GetSqlProcessID = 0
  End If

  rsUser.Close
  Set rsUser = Nothing
  
End Function

Public Function ReEstablishADOConnection() As Boolean

  Dim bDBLocked As Boolean
  Dim lngTimeOut As Long
  Dim bProgressVisible As Boolean
  Dim sErrorMsg As String

  lngTimeOut = 30

  If gobjProgress.Visible Then
    bProgressVisible = True
    gobjProgress.CloseProgress
  End If

  If COAMsgBox("A Windows network error has occured, and your connection to the database has been severed." & vbNewLine & vbNewLine _
      & "Would you like OpenHR to try reconnecting?", vbCritical + vbYesNo) = vbYes Then

    If Not datGeneral.Connect(gADOCon.ConnectionString, sErrorMsg, gsUserName, lngTimeOut, bDBLocked) Then
    
        COAMsgBox "OpenHR cannot establish a connection to the database." _
          & "This is likely a problem with your network connection. Please check your network connections, " & vbNewLine _
          & "or contact your IT system administrator." & vbNewLine & vbNewLine _
          & "OpenHR will now exit.", vbInformation
        ReEstablishADOConnection = False
        Stop
      
    Else
      If Not bDBLocked Then
        AuditAccess iDISCONNECTED, "Data"
        AuditAccess iRECONNECT, "Data"
        
        COAMsgBox "OpenHR has successfully logged back in." & vbNewLine & vbNewLine _
          & "If you experience this error on a regular basis it may point to a problem with your " _
          & "network environment. Please contact your IT system administrator.", vbInformation
          
        If bProgressVisible Then
          gobjProgress.OpenProgress
        End If
        
        ReEstablishADOConnection = True
      End If
    End If
  End If

End Function


Public Function ShowASCIIValues(strSourceString As String)
Dim iCount As Integer

  For iCount = 1 To Len(strSourceString)
    Debug.Print "ASC:" & Asc(Mid(strSourceString, iCount, 1)) & vbTab & vbTab & "<--" & vbTab & Chr(34) & Mid(strSourceString, iCount, 1) & Chr(34)
  Next
  
End Function

Private Function DisplayWarningToUser(userName As String, WarningType As WarningType, warningRefreshRate As Integer) As Boolean

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim bResult As Boolean

  On Error GoTo ErrorTrap
  
  ' Run the stored procedure to see if the given record has changed
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRUpdateWarningLog"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
                      
    Set pmADO = .CreateParameter("Username", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = userName
    
    Set pmADO = .CreateParameter("WarningType", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = WarningType
          
    Set pmADO = .CreateParameter("WarningRefreshRate", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = warningRefreshRate
          
    Set pmADO = .CreateParameter("WarnUser", adBoolean, adParamOutput)
    .Parameters.Append pmADO
          
    Set pmADO = Nothing

    cmADO.Execute
    bResult = CBool(.Parameters(3).Value)
    
  End With
               
TidyUpAndExit:
  DisplayWarningToUser = bResult
  Exit Function

ErrorTrap:
  bResult = False
  GoTo TidyUpAndExit

End Function
    
    

