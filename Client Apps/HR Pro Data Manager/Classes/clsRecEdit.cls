VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private datData As clsDataAccess

Private Sub Class_Initialize()

'  Set datData = New clsDataAccess
  Set datData = New HRProDataMgr.clsDataAccess

End Sub

Private Sub Class_Terminate()
  Set datData = Nothing

End Sub

Public Function NewLookupEntry(plngTableID As Long, psColumnName As String) As ADODB.Recordset
  ' Get the lookup table info for the given column.
  Dim sSQL As String
  Dim rsLookup As Recordset
    
  sSQL = "SELECT ASRSysColumns.lookupColumnID," & _
      " ASRSysColumns1.columnName," & _
      " ASRSysColumns.lookupTableID," & _
      " ASRSysTables.tableName" & _
    " FROM ASRSysColumns" & _
    " INNER JOIN ASRSysColumns ASRSysColumns1 ON ASRSysColumns.lookupColumnID = ASRSysColumns1.ColumnID" & _
    " INNER JOIN ASRSysTables ON ASRSysColumns.lookupTableID = ASRSysTables.TableID" & _
    " WHERE ASRSysColumns.tableID = " & Trim(Str(plngTableID)) & _
    " AND ASRSysColumns.columnName = '" & psColumnName & "'"
  Set rsLookup = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set NewLookupEntry = rsLookup

End Function

Public Function GetViewParentTable(lViewID As Long) As String

    Dim sSQL As String
    Dim rsTemp As Recordset
    
    sSQL = "SELECT ASRSysTables.TableName FROM ASRSysTables INNER JOIN ASRSysViews ON " & _
        "ASRSysTables.TableID = ASRSysViews.ViewTableID Where ASRSysViews.ViewID = " & lViewID
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
    GetViewParentTable = rsTemp(0).Value
    rsTemp.Close
    Set rsTemp = Nothing

End Function


Public Function GetScreen(plngScreenID As Long) As clsScreen
  ' Return the screen definition object for the given screen.
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim rsScreen As Recordset
  Dim rsPages As Recordset
  Dim objScreen As clsScreen
  Dim asTabCaptions() As String
  
  ' Instantiate the Screens collection if it does not already exist.
  If gcolScreens Is Nothing Then
    Set gcolScreens = New clsScreens
  End If
  
  ' If the given screen's definition has already been
  ' read then simply return it.
  For iLoop = 1 To gcolScreens.Count
    If gcolScreens.Item(iLoop).ScreenID = plngScreenID Then
      Set GetScreen = gcolScreens.Item(iLoop)
      Exit Function
    End If
  Next iLoop
  
  ' Read the given screen's definition and add it to the collection so that
  ' we do not have to hit the server next time the screen is called up.
  sSQL = "SELECT *" & _
    " FROM ASRSysScreens " & _
    " WHERE screenID=" & plngScreenID
  Set rsScreen = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsScreen
    If Not (.EOF And .BOF) Then
      Set objScreen = gcolScreens.Add(!ScreenID, !Name, !TableID, _
        !OrderID, !Height, !Width, !PictureID, !QuickEntry, _
        !FontName, !FontSize, !FontBold, !FontItalic, !FontStrikethru, !FontUnderline)
        
      ' Get the tab page collection for the screen.
      ReDim asTabCaptions(0)
      sSQL = "SELECT [Caption]" & _
        " FROM ASRSysPageCaptions " & _
        " WHERE screenID = " & Trim(Str(plngScreenID)) & _
        " ORDER BY pageIndexID"
      Set rsPages = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      With rsPages
        Do While Not .EOF
          iNextIndex = UBound(asTabCaptions) + 1
          ReDim Preserve asTabCaptions(iNextIndex)
          asTabCaptions(iNextIndex) = .Fields(0).Value
          
          .MoveNext
        Loop
        
        .Close
      End With
      Set rsPages = Nothing
      objScreen.TabCaptions = asTabCaptions
    End If
  
    .Close
  End With
  Set rsScreen = Nothing

  Set GetScreen = objScreen
  Set objScreen = Nothing
  
End Function

Public Function GetPages(lScreenID As Long) As ADODB.Recordset
    Dim sSQL As String
    Dim rsPages As Recordset
    
    sSQL = "SELECT * " & _
        "FROM ASRSysPageCaptions " & _
        "WHERE ScreenID = " & lScreenID & _
        " Order By PageIndexID"
    Set rsPages = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
    Set GetPages = rsPages

End Function

Public Function GetControls(plngScreenID As Long) As clsScreenControls
  ' Return the screen controls object defining the controls on the given screen.
  Dim iLoop As Integer
  Dim iControlCounter As Integer
  Dim sSQL As String
  Dim sKey As String
  Dim rsControls As Recordset
  Dim rsControlValues As Recordset
  Dim objScreenControl As clsScreenControl
  Dim objScreenControls As clsScreenControls
  Dim vControlValues As Variant
  
  ' Instantiate the Screen Controls collection if it does not already exist.
  If gcolScreenControls Is Nothing Then
    Set gcolScreenControls = New Collection
  End If

  ' If the given screen's controls collection has already been
  ' read then simply return it.
  For iLoop = 1 To gcolScreenControls.Count
    If gcolScreenControls.Item(iLoop).ScreenID = plngScreenID Then
      Set GetControls = gcolScreenControls.Item(iLoop)
      Exit Function
    End If
  Next iLoop

  ' Read the given screen controls definitions and add them to the collection so that
  ' we do not have to hit the server next time the screen is called up.
  Set objScreenControls = New clsScreenControls
  objScreenControls.ScreenID = plngScreenID
  iControlCounter = 0
  
  sSQL = "exec dbo.sp_ASRGetControlDetails " & plngScreenID
  Set rsControls = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsControls
    Do While Not .EOF
      ' Add the screen control definition to the collection.
      sKey = Trim(Str(iControlCounter))
      Set objScreenControl = objScreenControls.Add(sKey)
      iControlCounter = iControlCounter + 1
      
      objScreenControl.Key = sKey
      objScreenControl.PageNo = !PageNo
      objScreenControl.TableID = IIf(IsNull(!TableID), 0, !TableID)
      objScreenControl.ColumnID = IIf(IsNull(!ColumnID), 0, !ColumnID)
      objScreenControl.ControlType = !ControlType
      objScreenControl.TopCoord = !TopCoord
      objScreenControl.LeftCoord = !LeftCoord
      objScreenControl.Height = !Height
      objScreenControl.Width = !Width
      objScreenControl.Caption = IIf(IsNull(!Caption), "", !Caption)
      objScreenControl.BackColor = IIf(IsNull(!BackColor), 0, !BackColor)
      objScreenControl.ForeColor = IIf(IsNull(!ForeColor), 0, !ForeColor)
      objScreenControl.FontName = IIf(IsNull(!FontName), "", !FontName)
      objScreenControl.FontSize = IIf(IsNull(!FontSize), 0, !FontSize)
      objScreenControl.FontBold = IIf(IsNull(!FontBold), False, !FontBold)
      objScreenControl.FontItalic = IIf(IsNull(!FontItalic), False, !FontItalic)
      objScreenControl.FontStrikethru = IIf(IsNull(!FontStrikethru), False, !FontStrikethru)
      objScreenControl.FontUnderline = IIf(IsNull(!FontUnderline), False, !FontUnderline)
      objScreenControl.PictureID = IIf(IsNull(!PictureID), 0, !PictureID)
      objScreenControl.DisplayType = IIf(IsNull(!DisplayType), 0, !DisplayType)
      objScreenControl.TabIndex = !TabIndex
      objScreenControl.BorderStyle = IIf(IsNull(!BorderStyle), 0, !BorderStyle)
      objScreenControl.Alignment = IIf(IsNull(!Alignment), 0, !Alignment)
      objScreenControl.ColumnName = IIf(IsNull(!ColumnName), "", !ColumnName)
      objScreenControl.ColumnType = IIf(IsNull(!ColumnType), 0, !ColumnType)
      objScreenControl.DataType = IIf(IsNull(!DataType), 0, !DataType)
      objScreenControl.DefaultValue = IIf(IsNull(!DefaultValue), "", !DefaultValue)
      objScreenControl.Size = IIf(IsNull(!Size), 0, !Size)
      objScreenControl.Decimals = IIf(IsNull(!Decimals), 0, !Decimals)
      objScreenControl.LookupTableID = IIf(IsNull(!LookupTableID), 0, !LookupTableID)
      objScreenControl.LookupColumnID = IIf(IsNull(!LookupColumnID), 0, !LookupColumnID)
      objScreenControl.LookupFilterColumnID = IIf(IsNull(!LookupFilterColumnID), 0, !LookupFilterColumnID)
      objScreenControl.LookupFilterOperator = IIf(IsNull(!LookupFilterOperator), 0, !LookupFilterOperator)
      objScreenControl.LookupFilterValueID = IIf(IsNull(!LookupFilterValueID), 0, !LookupFilterValueID)
      objScreenControl.SpinnerMinimum = IIf(IsNull(!SpinnerMinimum), 0, !SpinnerMinimum)
      objScreenControl.SpinnerMaximum = IIf(IsNull(!SpinnerMaximum), 0, !SpinnerMaximum)
      objScreenControl.SpinnerIncrement = IIf(IsNull(!SpinnerIncrement), 0, !SpinnerIncrement)
      objScreenControl.Mandatory = IIf(IsNull(!Mandatory), False, !Mandatory)
      objScreenControl.UniqueCheck = IIf(IsNull(!UniqueCheckType), False, (!UniqueCheckType <> 0))
      objScreenControl.ConvertCase = IIf(IsNull(!ConvertCase), 0, !ConvertCase)
      objScreenControl.Mask = IIf(IsNull(!Mask), "", RTrim(!Mask))
      objScreenControl.BlankIfZero = IIf(IsNull(!BlankIfZero), False, !BlankIfZero)
      objScreenControl.Multiline = IIf(IsNull(!Multiline), False, !Multiline)
      objScreenControl.ColumnAlignment = IIf(IsNull(!ColAlignment), 0, !ColAlignment)
      objScreenControl.GotFocusExprID = IIf(IsNull(!GotFocusExprID), 0, !GotFocusExprID)
      objScreenControl.ReadOnly = IIf(IsNull(!ReadOnly), False, !ReadOnly)
      objScreenControl.StatusBarMessage = IIf(IsNull(!StatusBarMessage), "", !StatusBarMessage)
      objScreenControl.ErrorMessage = IIf(IsNull(!ErrorMessage), "", !ErrorMessage)
      objScreenControl.LinkTableID = IIf(IsNull(!LinkTableID), 0, !LinkTableID)
      objScreenControl.LinkViewID = IIf(IsNull(!LinkViewID), 0, !LinkViewID)
      objScreenControl.AFDEnabled = IIf(IsNull(!AFDEnabled), False, !AFDEnabled)
      objScreenControl.TableName = IIf(IsNull(!TableName), "", !TableName)
      objScreenControl.DfltValueExprID = IIf(IsNull(!DfltValueExprID), 0, !DfltValueExprID)
      objScreenControl.LinkOrderID = IIf(IsNull(!LinkOrderID), 0, !LinkOrderID)
      objScreenControl.Use1000Separator = IIf(IsNull(!Use1000Separator), False, !Use1000Separator)
      objScreenControl.TrimmingType = IIf(IsNull(!Trimming), giTRIMMING_NONE, !Trimming)
      objScreenControl.QAddressEnabled = IIf(IsNull(!QAddressEnabled), False, !QAddressEnabled)
      objScreenControl.OLEType = IIf(IsNull(!OLEType), 1, !OLEType)
      objScreenControl.EmbeddedEnabled = IIf(IsNull(!MaxOLESizeEnabled), False, !MaxOLESizeEnabled)
      objScreenControl.MaxOLESize = IIf(IsNull(!MaxOLESize), 0, !MaxOLESize)
      objScreenControl.AutoUpdateLookupValues = IIf(IsNull(!AutoUpdateLookupValues), False, !AutoUpdateLookupValues)
           
      ' Read the option group's control values.
      objScreenControl.ClearControlValues
      If (objScreenControl.ControlType = ctlRadio) Or _
        ((objScreenControl.ControlType = ctlCombo) And (objScreenControl.ColumnType <> colLookup)) Then
        Set rsControlValues = GetControlValues(!ColumnID)
        If Not rsControlValues.EOF And Not rsControlValues.BOF Then
          ' JPD 16/7/01 Do not limit the coontrol values to 100 items.
          'vControlValues = rsControlValues.GetRows(100)
          vControlValues = rsControlValues.GetRows
        Else
          ReDim vControlValues(0, 1)
          vControlValues(0, 0) = ""
          vControlValues(0, 1) = ""
        End If
        objScreenControl.ControlValues = vControlValues
        
        rsControlValues.Close
        Set rsControlValues = Nothing
      End If
      
      Set objScreenControl = Nothing
      
      .MoveNext
    Loop
    
    .Close
  End With
  Set rsControls = Nothing
    
  gcolScreenControls.Add objScreenControls, Trim(Str(plngScreenID))

  Set GetControls = objScreenControls
  Set objScreenControls = Nothing

End Function

Public Function GetViewDetails(plngViewID As Long)
  ' Return a recordset of the columns in the given view.
  Dim sSQL As String
  Dim rsViews As Recordset
    
  sSQL = "SELECT ASRSysViewColumns.inView, ASRSysColumns.columnName" & _
    " FROM ASRSysViewColumns" & _
    " INNER JOIN ASRSysColumns ON ASRSysViewColumns.columnID = ASRSysColumns.columnID " & _
    " WHERE ASRSysViewColumns.viewID = " & Trim(Str(plngViewID))
  Set rsViews = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetViewDetails = rsViews

End Function

Public Function GetControlValues(plngColumnID As Long) As ADODB.Recordset
  Dim sSQL As String
  Dim rsControlValues As Recordset
  
  sSQL = "SELECT value" & _
      " FROM ASRSysColumnControlValues" & _
      " WHERE columnID = " & plngColumnID & _
      " ORDER BY sequence"
  Set rsControlValues = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetControlValues = rsControlValues

End Function

Public Function GetHistoryTables(plngScreenID As Long) As ADODB.Recordset
  ' Return a recordset of the given screen's history tables.
  Dim sSQL As String
  Dim rsTables As Recordset
        
  sSQL = "SELECT DISTINCT ASRSysScreens1.tableID , ASRSysTables.tableName" & _
    " FROM ASRSysHistoryScreens" & _
    " INNER JOIN ASRSysScreens ON ASRSysHistoryScreens.parentScreenID = ASRSysScreens.screenID" & _
    " INNER JOIN ASRSysScreens ASRSysScreens1 ON ASRSysHistoryScreens.historyScreenID = ASRSysScreens1.screenID" & _
    " INNER JOIN ASRSysTables ON ASRSysScreens1.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysScreens.ScreenID = " & Trim(Str(plngScreenID))
  Set rsTables = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
  Set GetHistoryTables = rsTables

End Function

Public Function GetColumnDetails(lLookupTableID As Long, lLookupColumnID As Long) As ADODB.Recordset

    Dim sSQL As String
    Dim rsLookup As Recordset
    
    sSQL = "SELECT ASRSysColumns.ColumnName, ASRSysColumns.Mandatory, ASRSysTables.TableName " & _
        "FROM ASRSysColumns INNER JOIN ASRSysTables ON ASRSysTables.TableID=ASRSysColumns.TableID " & _
        "WHERE ASRSysColumns.TableID=" & lLookupTableID & " " & _
        "AND ASRSysColumns.ColumnID=" & lLookupColumnID
    Set rsLookup = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
    Set GetColumnDetails = rsLookup

End Function

'Public Function GetColumnRecordFromName(psTableName As String, psColumnName As String) As ADODB.Recordset
'
'    Dim sSQL As String
'    Dim rsColumn As Recordset
'    sSQL = "SELECT ASRSysColumns.*" & _
'        " FROM ASRSysColumns" & _
'        " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
'        " WHERE ASRSysColumns.columnName = '" & psColumnName & "'" & _
'        " AND ASRSysTables.tableName = '" & psTableName & "'"
'    Set rsColumn = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
'
'    Set GetColumnRecordFromName = rsColumn
'
'End Function


Public Function GetLookupDetails(plngTableID As Long, psColumnName As String) As ADODB.Recordset
  ' Return the lookup table and column id for the given column.
  Dim sSQL As String
  Dim rsTemp As Recordset
    
  sSQL = "SELECT lookupColumnID, lookupTableID" & _
    " FROM ASRSysColumns" & _
    " WHERE tableID = " & Trim(Str(plngTableID)) & _
    " AND columnName = '" & psColumnName & "'"
    
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetLookupDetails = rsTemp

End Function

Public Function GetColumnPrivilegeTables(lTableID As Long) As ADODB.Recordset

    Dim sSQL As String
    Dim rsColumns As Recordset
    
    sSQL = "Select ColumnName From ASRSysColumns Where TableID = " & lTableID & " AND ColumnName NOT LIKE 'ID%'"
    Set rsColumns = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
    Set GetColumnPrivilegeTables = rsColumns

End Function
'
'Public Function GetColumnPrivileges(sTableName As String, Optional sColName As String) As ADODB.Recordset
'
'    Dim sSQL As String
'    Dim rsPriv As Recordset
'
''    sSql = "Select COLUMN_NAME, PRIVILEGE_TYPE From vew_ASRColumnPrivileges Where " & _
''        "GRANTEE = '" & gsUserGroup & "' And TABLE_NAME = '" & sTableName & "'"
''    If Len(sColName) > 0 Then
''        sSql = sSql & " And COLUMN_NAME = '" & sColName & "'"
''    End If
''    sSql = sSql & " And COLUMN_NAME <> 'TIMESTAMP' Order By COLUMN_NAME"
'
'    sSQL = "exec sp_ASRGetColumnPrivileges '" & gsUserGroup & "', '" & sTableName & "', '" & sColName & "'"
'
'
'    Set rsPriv = datData.OpenRecordset(sSQL, adOpenKeyset, adLockReadOnly)
'    Set GetColumnPrivileges = rsPriv
'
'End Function

Public Function GetLookupPrivileges(sTableName As String)

    Dim sSQL As String
    Dim rsPriv As Recordset
    
    sSQL = "SELECT ASRSysTables.TableName, ASRSysColumns1.ColumnName FROM ASRSysColumns INNER JOIN " & _
        "ASRSysTables ON ASRSysColumns.lookupTableID = ASRSysTables.TableID INNER JOIN " & _
        "ASRSysColumns ASRSysColumns1 ON ASRSysColumns.lookupColumnID = ASRSysColumns1.ColumnID " & _
        "WHERE ASRSysColumns.columnName = '" & sTableName & "'"
    Set rsPriv = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Set GetLookupPrivileges = rsPriv

End Function

Public Function GetColumnNamesFromName(psTableViewName As String) As ADODB.Recordset
  ' Return a recordset of the the columns in the given table/view.
  Dim sSQL As String
  Dim rsColumns As Recordset
    
  sSQL = "SELECT ASRSysColumns.*" & _
    " FROM ASRSysColumns" & _
    " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysTables.tableName = '" & psTableViewName & "'" & _
    " UNION " & _
    " SELECT ASRSysColumns.*" & _
    " FROM ASRSysColumns" & _
    " INNER JOIN ASRSysViews ON ASRSysColumns.tableID = ASRSysviews.viewtableID" & _
    " INNER JOIN ASRSysViewColumns ON ASRSysColumns.columnID = ASRSysViewColumns.columnID" & _
    " WHERE ASRSysViews.viewName = '" & psTableViewName & "'" & _
    " AND ASRSysViewColumns.viewID = ASRSysViews.viewID" & _
    " AND asrSysViewColumns.inView = 1"

  Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetColumnNamesFromName = rsColumns

End Function



'Public Function CanSelect(sTableName As String, sColName As String) As Boolean
'
'    Dim rsSelect As Recordset
'
'    Set rsSelect = GetColumnPrivileges(sTableName, sColName)
'    Do While Not rsSelect.EOF
'        If rsSelect!PRIVILEGE_TYPE = "SELECT" Then
'            CanSelect = True
'            Exit Function
'        End If
'        rsSelect.MoveNext
'    Loop
'    CanSelect = False
'
'End Function

Public Function GetLinkID(sTable As String, lLinkTableID As Long, lID As Long) As Long

    Dim sSQL As String
    Dim rsTemp As Recordset
    
    sSQL = "Select ID_" & lLinkTableID & " From " & sTable & " Where ID = " & lID
    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
    GetLinkID = IIf(IsNull(rsTemp(0).Value), 0, rsTemp(0).Value)
    rsTemp.Close
    Set rsTemp = Nothing

End Function

Public Function CanInsert(sTableName As String) As Boolean
  Dim sSQL As String
  Dim rsTemp As Recordset
    
  sSQL = "EXEC sp_table_privileges_rowset '" & sTableName & "', null, null, '" & datGeneral.UserNameForSQL & "'"
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
  Do While Not rsTemp.EOF
    If rsTemp!PRIVILEGE_TYPE = "INSERT" Then
      CanInsert = True
      GoTo Exit_Out
    End If
    
    rsTemp.MoveNext
  Loop
  CanInsert = False
    
Exit_Out:
  rsTemp.Close
  Set rsTemp = Nothing
    
End Function
'
'Public Function GetLinkColumns(plngScreenID As Long) As ADODB.Recordset
'  Dim sSQL As String
'  Dim rsTemp As Recordset
'
'  sSQL = "exec dbo.sp_ASRGetLinkColumns " & plngScreenID & ", '" & gsUserGroup & "'"
'  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'
'  Set GetLinkColumns = rsTemp
'
'End Function
'
'Public Function GetLinkViewColumns(lViewID As Long, lScreenID As Long) As ADODB.Recordset
'
'    Dim sSQL As String
'    Dim rsTemp As Recordset
'
'    sSQL = "exec dbo.sp_ASRGetViewLinkColumnPrivileges " & lViewID & ", " & lScreenID & ", '" & gsUserGroup & "'"
'    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'    Set GetLinkViewColumns = rsTemp
'
'End Function
'
'Public Function GetAllLinkViewColumns(lScreenID As Long) As ADODB.Recordset
'
'    Dim sSQL As String
'    Dim rsTemp As Recordset
'
'    sSQL = "exec dbo.sp_ASRGetViewLinkColumns " & lScreenID
'    Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'    Set GetAllLinkViewColumns = rsTemp
'
'End Function



