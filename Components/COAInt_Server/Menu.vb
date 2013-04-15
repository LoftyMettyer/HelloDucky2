Option Strict Off
Option Explicit On
Public Class Menu

  Private mclsData As clsDataAccess

  Public WriteOnly Property Connection() As Object
    Set(ByVal Value As Object)

      ' Connection object passed in from the asp page
      Dim sGroupName As String
      Dim sSQL As String
      Dim rsUser As ADODB.Recordset
      Dim datData As clsDataAccess

      ' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
      If ASRDEVELOPMENT Then
        gADOCon = New ADODB.Connection
        'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gADOCon.Open(Value)
        CreateASRDev_SysProtects(gADOCon)
      Else
        gADOCon = Value
      End If

      '  Set datData = New clsDataAccess
      '  sSQL = "exec sp_helpuser '" & gsUsername & "'"
      '  Set rsUser = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      '  If rsUser!GroupName = "db_owner" Then
      '    rsUser.MoveNext
      '  End If
      '  sGroupName = rsUser!GroupName
      '  rsUser.Close
      '  Set rsUser = Nothing
      '  Set datData = Nothing

      '  If sGroupName <> gsUserGroup Then
      'JPD 20031006 Yes, do drop the tables & columns collections as we want to be sure that
      'the current users settings are read. HDA encountered a problem where they weren't.
      ' JPD20030313 Do not drop the tables & columns collections as they can be reused.
      'UPGRADE_NOTE: Object gcoTablePrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      gcoTablePrivileges = Nothing
      'UPGRADE_NOTE: Object gcolColumnPrivilegesCollection may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      gcolColumnPrivilegesCollection = Nothing
      '  End If

      '  gsUserGroup = sGroupName

      SetupTablesCollection()

    End Set
  End Property

  Public WriteOnly Property Username() As String
    Set(ByVal value As String)

      ' Username passed in from the asp page
      gsUsername = value

    End Set
  End Property

  Public Function GetHistoryScreens() As Object
    ' Return an array of information that can be used to format the History tables menu for the current user.
    ' The recordset contains a row for each primary table in the HR Pro database.

    On Error GoTo ErrorTrap

    Dim sSQL As String
    Dim rsTableScreens As ADODB.Recordset
    Dim avTableInfo(,) As Object
    Dim iNextIndex As Short

    ReDim avTableInfo(5, 0)
    ' Index 1 = parent screen ID
    ' Index 2 = child table ID
    ' Index 3 = child table name
    ' Index 4 = child table screen ID
    ' Index 5 = child table screen name

    sSQL = "SELECT ASRSysTables.tableName AS [childTableName]," & " childScreens.tableID AS [childTableID]," & " childScreens.screenID AS [childScreenID]," & " childScreens.name AS [childScreenName]," & " parentScreen.screenid AS [parentScreenID]" & " FROM ASRSysScreens parentScreen" & " INNER JOIN ASRSysHistoryScreens" & "   ON parentScreen.screenID = ASRSysHistoryScreens.parentScreenID" & " INNER JOIN ASRSysScreens childScreens" & "   ON ASRSysHistoryScreens.historyScreenID = childScreens.screenID" & " INNER JOIN ASRSysTables" & "   ON childScreens.tableID = ASRSysTables.tableID" & " WHERE childScreens.quickEntry = 0" & " ORDER BY parentScreen.screenid," & "   ASRSysTables.tableName DESC," & "   childScreens.Name DESC"

    rsTableScreens = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

    Do While Not rsTableScreens.EOF
      If gcoTablePrivileges.Item(rsTableScreens.Fields("childTableName").Value).AllowSelect Then
        iNextIndex = UBound(avTableInfo, 2) + 1
        ReDim Preserve avTableInfo(5, iNextIndex)

        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        avTableInfo(1, iNextIndex) = rsTableScreens.Fields("parentScreenID").Value
        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        avTableInfo(2, iNextIndex) = rsTableScreens.Fields("childTableID").Value
        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        avTableInfo(3, iNextIndex) = rsTableScreens.Fields("childTableName").Value
        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(4, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        avTableInfo(4, iNextIndex) = rsTableScreens.Fields("childScreenID").Value
        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(5, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        avTableInfo(5, iNextIndex) = Replace(rsTableScreens.Fields("childScreenName").Value, "&", "&&")
      End If

      rsTableScreens.MoveNext()
    Loop

    rsTableScreens.Close()
    'UPGRADE_NOTE: Object rsTableScreens may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTableScreens = Nothing

TidyUpAndExit:
    'UPGRADE_WARNING: Couldn't resolve default property of object GetHistoryScreens. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    GetHistoryScreens = VB6.CopyArray(avTableInfo)
    Exit Function

ErrorTrap:

  End Function

  Public Function GetPrimaryTableMenu() As Object

    On Error GoTo ErrorTrap

    ' Return an array of information that can be used to format the main Database menu for the current user.
    ' The recordset contains a row for each primary table in the HR Pro database.
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
    Dim iNextIndex As Short
    Dim iTotalViewScreenCount As Short
    Dim sSQL As String
    Dim sViewList As String
    Dim rsViews As ADODB.Recordset
    Dim rsViewScreen As ADODB.Recordset
    Dim rsTables As ADODB.Recordset
    Dim rsTableScreen As ADODB.Recordset
    Dim avTableInfo(,) As Object
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
    sSQL = "SELECT ASRSysTables.tableID," & " ASRSysTables.tableName," & " COUNT(DISTINCT ASRSysScreens.name) AS tableScreenCount" & " FROM ASRSysTables" & " INNER JOIN ASRSysScreens" & " ON ASRSysTables.tableID = ASRSysScreens.tableID" & " AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))" & " AND ((ASRSysScreens.quickEntry IS null) OR (ASRSysScreens.quickEntry = 0))" & " GROUP BY ASRSysTables.tableID," & " ASRSysTables.tableName," & " ASRSysTables.tableType" & " HAVING ASRSysTables.tableType = 1" & " ORDER BY ASRSysTables.tableName DESC"
    rsTables = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    Do While Not rsTables.EOF
      ' Initialise an entry into our array of table info for each primary table.
      iNextIndex = UBound(avTableInfo, 2) + 1
      ReDim Preserve avTableInfo(12, iNextIndex)

      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(1, iNextIndex) = rsTables.Fields("TableID").Value
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(2, iNextIndex) = rsTables.Fields("TableName").Value
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(3, iNextIndex) = rsTables.Fields("tableScreenCount").Value
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(4, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(4, iNextIndex) = 0
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(5, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(5, iNextIndex) = gcoTablePrivileges.Item(rsTables.Fields("TableName").Value).AllowSelect
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(6, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(6, iNextIndex) = 0
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(7, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(7, iNextIndex) = 0
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(8, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(8, iNextIndex) = ""
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(9, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(9, iNextIndex) = 0
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(10, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(10, iNextIndex) = 0
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(11, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(11, iNextIndex) = 0
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(12, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avTableInfo(12, iNextIndex) = 0

      iTotalViewScreenCount = 0

      ' Create a list of the current user's permitted views on the current table.
      sViewList = "0"
      For Each objTableView In gcoTablePrivileges.Collection
        If Not (objTableView.IsTable) And (objTableView.TableID = rsTables.Fields("TableID").Value) And (objTableView.AllowSelect) Then

          sViewList = sViewList & ", " & Trim(Str(objTableView.ViewID))
        End If
      Next objTableView
      'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      objTableView = Nothing

      ' Get view information for the current table.
      sSQL = "SELECT ASRSysViews.viewName, COUNT (ASRSysViewScreens.ScreenID) AS viewScreenCount" & " FROM ASRSysViews" & " LEFT OUTER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID" & " GROUP BY ASRSysViews.viewName, ASRSysViews.viewID" & " HAVING ASRSysViews.viewID IN (" & sViewList & ")"
      rsViews = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
      Do While Not rsViews.EOF
        If rsViews.Fields("viewScreenCount").Value > 0 Then
          'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(6, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          avTableInfo(6, iNextIndex) = avTableInfo(6, iNextIndex) + 1
          iTotalViewScreenCount = iTotalViewScreenCount + rsViews.Fields("viewScreenCount").Value
        End If

        rsViews.MoveNext()
      Loop
      rsViews.Close()
      'UPGRADE_NOTE: Object rsViews may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      rsViews = Nothing

      ' Get view screen info if required.
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(6, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If avTableInfo(6, iNextIndex) > 0 Then
        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(9, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        avTableInfo(9, iNextIndex) = iTotalViewScreenCount

        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(5, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (iTotalViewScreenCount = 1) And ((avTableInfo(5, iNextIndex) = False) Or (avTableInfo(3, iNextIndex) = 0)) Then

          sSQL = "SELECT ASRSysViews.viewID, ASRSysViews.viewName, ASRSysViewScreens.screenID, ASRSysScreens.pictureID" & " FROM ASRSysViews" & " INNER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID" & " INNER JOIN ASRSysScreens ON ASRSysViewScreens.screenID = ASRSysScreens.screenID" & " WHERE ASRSysViews.viewid IN (" & sViewList & ")"
          rsViewScreen = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
          If Not (rsViewScreen.EOF And rsViewScreen.BOF) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(7, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            avTableInfo(7, iNextIndex) = rsViewScreen.Fields("ViewID").Value
            'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(8, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            avTableInfo(8, iNextIndex) = rsViewScreen.Fields("ViewName").Value
            'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(10, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            avTableInfo(10, iNextIndex) = rsViewScreen.Fields("ScreenID").Value
            'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(12, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            avTableInfo(12, iNextIndex) = rsViewScreen.Fields("pictureID").Value
          End If
          rsViewScreen.Close()
          'UPGRADE_NOTE: Object rsViewScreen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
          rsViewScreen = Nothing
        End If
      End If

      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(5, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If (avTableInfo(3, iNextIndex) = 1) And (iTotalViewScreenCount = 0) And (avTableInfo(5, iNextIndex) = True) Then
        sSQL = "SELECT ASRSysScreens.screenID," & " ASRSysScreens.pictureID" & " FROM ASRSysScreens" & " WHERE ASRSysScreens.tableID = " & Trim(Str(rsTables.Fields("TableID").Value)) & "   AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))"

        rsTableScreen = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If Not (rsTableScreen.EOF And rsTableScreen.BOF) Then
          'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(4, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          avTableInfo(4, iNextIndex) = rsTableScreen.Fields("ScreenID").Value
          'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(11, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          avTableInfo(11, iNextIndex) = rsTableScreen.Fields("pictureID").Value
        End If
        rsTableScreen.Close()
        'UPGRADE_NOTE: Object rsTableScreen may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsTableScreen = Nothing
      End If

      rsTables.MoveNext()
    Loop
    rsTables.Close()
    'UPGRADE_NOTE: Object rsTables may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTables = Nothing
    'UPGRADE_WARNING: Couldn't resolve default property of object GetPrimaryTableMenu. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    GetPrimaryTableMenu = VB6.CopyArray(avTableInfo)

TidyUpAndExit:
    Exit Function
ErrorTrap:

  End Function

  Public Function GetPrimaryTableSubMenu(ByVal plngTableID As Integer) As Object
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
    Dim iNextIndex As Short
    Dim sSQL As String
    Dim sViewList As String
    Dim rsScreens As ADODB.Recordset
    Dim avScreenInfo(,) As Object
    Dim objTableView As CTablePrivilege
    Dim sTableName As String

    ' Create an array with records for each screen for each permitted view on the primary table.
    ReDim avScreenInfo(5, 0)

    ' Create a list of the current user's permitted views on the current table.
    sViewList = "0"
    For Each objTableView In gcoTablePrivileges.Collection
      If Not (objTableView.IsTable) And (objTableView.TableID = plngTableID) And (objTableView.AllowSelect) Then

        sViewList = sViewList & ", " & Trim(Str(objTableView.ViewID))
      End If

      If (objTableView.IsTable) And (objTableView.TableID = plngTableID) Then

        sTableName = objTableView.TableName
      End If
    Next objTableView
    'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objTableView = Nothing

    sSQL = "SELECT ASRSysViewScreens.screenID," & " ASRSysScreens.name," & " ASRSysViews.viewID," & " ASRSysViews.ViewName," & " ASRSysScreens.PictureID" & " FROM ASRSysViews" & " INNER JOIN ASRSysViewScreens ON ASRSysViews.viewID = ASRSysViewScreens.viewID" & " INNER JOIN ASRSysScreens ON ASRSysViewScreens.screenID = ASRSysScreens.screenID" & " WHERE ASRSysViews.viewID IN (" & sViewList & ")"

    If gcoTablePrivileges.Item(sTableName).AllowSelect Then
      ' The current user does have SELECT permission on the given table, so populate the array
      ' table with records for each screen associated with the primary table.
      sSQL = sSQL & " UNION" & " SELECT ASRSysScreens.screenID," & " ASRSysScreens.Name," & " 0 AS viewID," & " '' AS viewName," & " ASRSysScreens.pictureID" & " FROM ASRSysScreens" & " WHERE (ASRSysScreens.tableID = " & Trim(Str(plngTableID)) & ")" & " AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))" & " AND ((ASRSysScreens.quickEntry IS null) OR (ASRSysScreens.quickEntry = 0))"
    End If

    sSQL = sSQL & " ORDER BY ASRSysScreens.name DESC, viewName DESC"

    rsScreens = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
    Do While Not rsScreens.EOF
      iNextIndex = UBound(avScreenInfo, 2) + 1
      ReDim Preserve avScreenInfo(5, iNextIndex)
      'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avScreenInfo(1, iNextIndex) = rsScreens.Fields("ScreenID").Value
      'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avScreenInfo(2, iNextIndex) = Replace(rsScreens.Fields("Name").Value, "&", "&&")
      'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avScreenInfo(3, iNextIndex) = rsScreens.Fields("ViewID").Value
      'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(4, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avScreenInfo(4, iNextIndex) = rsScreens.Fields("ViewName").Value
      'UPGRADE_WARNING: Couldn't resolve default property of object avScreenInfo(5, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      avScreenInfo(5, iNextIndex) = rsScreens.Fields("pictureID").Value

      rsScreens.MoveNext()
    Loop
    rsScreens.Close()
    'UPGRADE_NOTE: Object rsScreens may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsScreens = Nothing

    'UPGRADE_WARNING: Couldn't resolve default property of object GetPrimaryTableSubMenu. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    GetPrimaryTableSubMenu = VB6.CopyArray(avScreenInfo)

  End Function

  Public Function GetTableScreens() As Object
    ' Return an array of information that can be used to format the Lookup tables menu for the current user.
    ' The recordset contains a row for each primary table in the HR Pro database.

    On Error GoTo ErrorTrap

    Dim sSQL As String
    Dim rsTableScreens As ADODB.Recordset
    Dim rsPermissions As ADODB.Recordset
    Dim avTableInfo(,) As Object
    Dim iNextIndex As Short

    ReDim avTableInfo(3, 0)
    ' Index 1 = table ID
    ' Index 2 = table name
    ' Index 3 = table screen ID

    sSQL = "SELECT ASRSysTables.tableID, ASRSysTables.tableName, ASRSysScreens.screenID" & " FROM ASRSysTables" & " INNER JOIN ASRSysScreens ON ASRSysTables.tableID = ASRSysScreens.tableID" & " WHERE ASRSysTables.tableType = " & Trim(Str(Declarations.TableTypes.tabLookup)) & " AND ((ASRSysScreens.ssIntranet IS null) OR (ASRSysScreens.ssIntranet = 0))" & " AND ((ASRSysScreens.quickEntry IS null) OR (ASRSysScreens.quickEntry = 0))" & " ORDER BY ASRSysTables.tableName DESC"

    rsTableScreens = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

    Do While Not rsTableScreens.EOF
      'Fault 11847 - The sysusers table will not always contain the 'HR Pro' group id in the
      ' gid column i.e. it could contain ASRSysGroup, Public etc. therefore the sysusers table
      ' can not be joined to the ASRSysViewPermissions table on the group name.
      '    sSQL = "SELECT COUNT(*) AS [result]" & _
      ''      " FROM ASRSysViewMenuPermissions" & _
      ''      " INNER JOIN sysusers b ON ASRSysViewMenuPermissions.groupName = b.name" & _
      ''      " INNER JOIN sysusers a on b.uid = a.gid" & _
      ''      "   AND a.name = current_user" & _
      ''      " WHERE ASRSysViewMenuPermissions.tableName = '" & rsTableScreens!TableName & "'" & _
      ''      "   AND ASRSysViewMenuPermissions.hideFromMenu = 1"
      sSQL = "SELECT COUNT(*) AS [result]" & " FROM ASRSysViewMenuPermissions" & " WHERE ASRSysViewMenuPermissions.tableName = '" & rsTableScreens.Fields("TableName").Value & "'" & "   AND ASRSysViewMenuPermissions.groupName = '" & gsUserGroup & "'" & "   AND ASRSysViewMenuPermissions.hideFromMenu = 1"

      rsPermissions = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

      If rsPermissions.Fields("Result").Value = 0 Then
        iNextIndex = UBound(avTableInfo, 2) + 1
        ReDim Preserve avTableInfo(3, iNextIndex)

        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        avTableInfo(1, iNextIndex) = rsTableScreens.Fields("TableID").Value
        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        avTableInfo(2, iNextIndex) = rsTableScreens.Fields("TableName").Value
        'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        avTableInfo(3, iNextIndex) = rsTableScreens.Fields("ScreenID").Value
      End If

      rsPermissions.Close()
      'UPGRADE_NOTE: Object rsPermissions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      rsPermissions = Nothing

      rsTableScreens.MoveNext()
    Loop

    'UPGRADE_NOTE: Object rsTableScreens may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTableScreens = Nothing

TidyUpAndExit:

    'UPGRADE_WARNING: Couldn't resolve default property of object GetTableScreens. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    GetTableScreens = VB6.CopyArray(avTableInfo)
    Exit Function

ErrorTrap:
    'Open "c:\temp\test.txt" For Append As #99
    'Print #99, "GetTableScreens  " & sSQL
    'Close #99

  End Function

  Public Function GetQuickEntryScreens() As Object
    Dim sSQL As String
    Dim rsScreens As ADODB.Recordset
    Dim avTableInfo(,) As Object
    Dim iNextIndex As Short

    ReDim avTableInfo(3, 0)
    ' Index 1 = table ID
    ' Index 2 = screen name
    ' Index 3 = table screen ID

    sSQL = "SELECT ASRSysScreens.screenID, ASRSysScreens.name, " & " ASRSysTables.tableName, ASRSysTables.tableID" & " FROM ASRSysScreens" & " INNER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID" & " WHERE ASRSysScreens.quickEntry = 1"

    rsScreens = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    Do While Not rsScreens.EOF
      'First see if we have privileges to see this table
      If gcoTablePrivileges.Item(rsScreens.Fields("TableName").Value).AllowSelect Then

        ' Check that the current user has 'select' permission on at least one parent table,
        ' or at least one view of one parent table referenced by the quick entry screen.
        If ViewQuickEntry(rsScreens.Fields("ScreenID").Value) Then
          iNextIndex = UBound(avTableInfo, 2) + 1
          ReDim Preserve avTableInfo(3, iNextIndex)

          'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(1, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          avTableInfo(1, iNextIndex) = rsScreens.Fields("TableID").Value
          'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(2, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          avTableInfo(2, iNextIndex) = Replace(rsScreens.Fields("Name").Value, "&", "&&")
          'UPGRADE_WARNING: Couldn't resolve default property of object avTableInfo(3, iNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          avTableInfo(3, iNextIndex) = rsScreens.Fields("ScreenID").Value
        End If
      End If
      rsScreens.MoveNext()
    Loop

    'UPGRADE_NOTE: Object rsScreens may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsScreens = Nothing

    'UPGRADE_WARNING: Couldn't resolve default property of object GetQuickEntryScreens. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    GetQuickEntryScreens = VB6.CopyArray(avTableInfo)

  End Function

  Private Function ViewQuickEntry(ByVal plngScreenID As Integer) As Boolean
    ' Return TRUE if the current user can see at least one parent table (or view of a parent table)
    ' of given quick view screen.
    On Error GoTo ErrorTrap

    Dim fCanView As Boolean
    Dim rsTables As ADODB.Recordset
    Dim rsViews As ADODB.Recordset

    fCanView = False

    ' Get the list of parent tables used in the quick entry screen.
    rsTables = GetQuickEntryTables(plngScreenID)

    With rsTables
      If (.EOF And .BOF) Then
        fCanView = True
      End If

      ' Loop through parent tables, seeing if we have select permissions on these tables.
      Do While (Not .EOF) And (Not fCanView)
        ' Check if the current user has 'select' permission on the given table.
        If gcoTablePrivileges.Item(.Fields("TableName").Value).AllowSelect Then
          fCanView = True
        Else
          ' No select permissions, can we use a view instead ???
          rsViews = GetQuickEntryViews(.Fields("TableID").Value)

          'Loop through the views, and see if we have permission on these
          Do While (Not rsViews.EOF) And (Not fCanView)
            If gcoTablePrivileges.Item(rsViews.Fields("ViewName")).AllowSelect Then
              'We have a view we can use, let's get outta here
              fCanView = True
            End If

            rsViews.MoveNext()
          Loop
          rsViews.Close()
          'UPGRADE_NOTE: Object rsViews may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
          rsViews = Nothing
        End If

        .MoveNext()
      Loop
      .Close()
    End With
    'UPGRADE_NOTE: Object rsTables may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTables = Nothing

    ViewQuickEntry = fCanView
    Exit Function

ErrorTrap:
    If Err.Number = 457 Then
      Resume Next
    End If

  End Function

  Public Function GetQuickEntryTables(ByVal plngScreenID As Integer) As ADODB.Recordset
    ' Return a recordset of all the table id's of the controls which aren't related to the base table
    ' in the given screen.
    Dim sSQL As String
    Dim rsTemp As ADODB.Recordset

    sSQL = "exec sp_ASRGetQuickEntryTables " & plngScreenID
    rsTemp = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
    GetQuickEntryTables = rsTemp

  End Function

  Public Function GetQuickEntryViews(ByVal plngTableID As Integer) As ADODB.Recordset
    ' Return a recordset of the user defined views on the given table.
    Dim sSQL As String
    Dim rsViews As ADODB.Recordset

    sSQL = "SELECT viewID, viewName" & " FROM ASRSysViews" & " WHERE viewTableID = " & Trim(Str(plngTableID))
    rsViews = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
    GetQuickEntryViews = rsViews

  End Function

  'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
  Private Sub Class_Initialize_Renamed()
    mclsData = New clsDataAccess
  End Sub

  Public Sub New()
    MyBase.New()
    Class_Initialize_Renamed()
  End Sub

  'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
  Private Sub Class_Terminate_Renamed()
    'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mclsData = Nothing

  End Sub

  Protected Overrides Sub Finalize()
    Class_Terminate_Renamed()
    MyBase.Finalize()
  End Sub
End Class