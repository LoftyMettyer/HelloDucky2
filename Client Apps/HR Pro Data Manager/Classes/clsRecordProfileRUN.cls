VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordProfileRUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngRecordProfileID As Long
Private msRecordProfileName As String
Private mstrRecordProfileDescription As String
Private mbDefinitionOwner As Boolean
Private mobjDefinition As clsRecordProfileTabDtls
Private msTempTableName As String
Private mfPrintFilterHeader As Boolean

Private mlngSingleRecordID As Long

'New Default Output Variables
Private mlngOutputFormat As Long
Private mblnOutputScreen As Boolean
Private mblnOutputPrinter As Boolean
Private mstrOutputPrinterName As String
Private mblnOutputSave As Boolean
Private mlngOutputSaveExisting As Long
'Private mlngOutputSaveFormat As Long
Private mblnOutputEmail As Boolean
Private mlngOutputEmailAddr As Long
Private mstrOutputEmailSubject As String
Private mstrOutputEmailAttachAs As String
'Private mlngOutputEmailFileFormat As Long
Private mstrOutputFileName As String
Private mblnPreviewOnScreen As Boolean
Private mstrOutputTitlePage As String
Private mstrOutputReportPackTitle As String
Private mstrOutputOverrideFilter As String
Private mblnOutputTOC As Boolean
Private mblnOutputCoverSheet As Boolean
Private mlngOverrideFilterID As Long
Private mblnChkPicklistFilter As Boolean 'might not need

Private mblnIndentRelatedTables As String
Private mblnSuppressEmptyRelatedTableTitles As String
Private mblnSuppressTableRelationshipTitles As String

Private miIndex As Integer

Private msErrorString As String

' String to hold the SQL statement
Private msSQL As String

' Data access class
Private mclsData As clsDataAccess
Private mclsGeneral As clsGeneral
Private mrsResults As ADODB.Recordset

' Instance of the previewform
Private mfrmOutput As frmRecordProfilePreview

' TableViewsGuff
Private mavTableViews() As Variant
' Column 1 = 0 if this row is for a table, 1 if it is for a view
' Column 2 = table/view ID
' Column 3 = table ID to which this table/view is related
' Column 4 = alias (if required)

'Batch Job Mode ?
'Private gblnBatchMode As Boolean

'Has the user cancelled the report ?
Private mblnUserCancelled As Boolean

'Does the record profile generate no records ?
Private mfNoRecords As Boolean

Private mblnInvalidPicklistFilter As Boolean

Private Const sTYPECODE_HEADING = "H"
Private Const sTYPECODE_SEPARATOR = "S"
Private Const sTYPECODE_COLUMN = "C"
Private Const sTYPECODE_ID = "I"

Private mvarUDFsRequired() As String

Private Sub GetRecordProfileTableOrder(psRelationship As String, plngTableID As Long, piGeneration As Integer)
  Dim iNextIndex As Integer
  Dim iLoop As Integer
  Dim sSQL As String
  Dim rsTables As ADODB.Recordset
  Dim objTable As clsRecordProfileTabDtl
  
  If psRelationship = "CHILD" Then
    sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
      " FROM ASRSysTables " & _
      " INNER JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.childID" & _
      " WHERE ASRSysRelations.parentID = " & Trim(Str(plngTableID))
  Else
    sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
      " FROM ASRSysTables " & _
      " INNER JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.parentID" & _
      " WHERE ASRSysRelations.childID = " & Trim(Str(plngTableID))
  End If
    
  Set rsTables = mclsData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  If Not (rsTables.BOF And rsTables.EOF) Then
  
    For iNextIndex = 2 To mobjDefinition.Count
      rsTables.MoveFirst
      
      Do While Not rsTables.EOF
        Set objTable = mobjDefinition.ItemByPosition(iNextIndex)
        
        If rsTables!TableID = objTable.TableID Then
          ' Remember which related tables have children in the record profile.
          If psRelationship = "CHILD" Then
            mobjDefinition.Item(CStr(plngTableID)).HasChildren = True
            objTable.RelatedTableID = plngTableID
            objTable.Relationship = psRelationship
            objTable.Generation = piGeneration
          Else
            objTable.HasChildren = False
            objTable.RelatedTableID = plngTableID
            objTable.Relationship = psRelationship
            objTable.Generation = piGeneration
          End If
                             
          objTable.Sequence = miIndex
          miIndex = miIndex + 1
          
          GetRecordProfileTableOrder psRelationship, objTable.TableID, piGeneration + 1
          Exit Do
        End If
      
        Set objTable = Nothing
        
        rsTables.MoveNext
      Loop
    Next iNextIndex
  End If
  
  rsTables.Close
  Set rsTables = Nothing

End Sub

Public Function RunRecordProfile(Optional plngRecordID As Long) As Boolean
  ' Purpose : This function is called from frmDefsel and Batch Jobs.
  '           It is the main function which runs the Record Profile.
  On Error GoTo RunRecordProfile_ERROR

  Dim fOK As Boolean

  fOK = True

  ReDim mvarUDFsRequired(0)

  Screen.MousePointer = vbHourglass
  mlngSingleRecordID = plngRecordID

  If fOK Then fOK = GetRecordProfileTables

  gobjEventLog.AddHeader eltRecordProfile, msRecordProfileName

  If fOK Then
    With gobjProgress
      '.AviFile = App.Path & "\videos\report.avi"
      .AVI = dbText
      .MainCaption = "Record Profile"
      If Not gblnBatchMode Then
        .NumberOfBars = 1
        .Caption = "Record Profile"
        .Time = False
        .Cancel = True
        .Bar1Caption = "Record Profile : " & msRecordProfileName
        .OpenProgress
      Else
        .Bar2Caption = "Record Profile : " & msRecordProfileName
      End If
    End With
  End If

  If fOK Then fOK = GetRecordProfileColumns
  
  'JPD 20030529 Fault 5865
  If fOK Then fOK = GenerateFastSQL
  
  If fOK Then fOK = ExecuteSql
  If fOK Then fOK = CheckRecordSet
  If fOK Then fOK = InitialisePreview

  If fOK Then
    mblnUserCancelled = mfrmOutput.UserCancelled

    If Not mblnUserCancelled Then
      If gblnBatchMode Or Not mblnPreviewOnScreen Then
        fOK = mfrmOutput.ShowResults
        If fOK Then
          fOK = mfrmOutput.OutputReport(False)
        End If
        If Not fOK Then msErrorString = mfrmOutput.ErrorMessage
        gobjProgress.UpdateProgress True
      ElseIf mblnPreviewOnScreen Then
        fOK = mfrmOutput.ShowResults
        If Not fOK Then msErrorString = mfrmOutput.ErrorMessage
        
        If fOK And Not mblnUserCancelled Then
          gobjProgress.CloseProgress
          Screen.MousePointer = vbDefault
          If fOK Then mfrmOutput.Show vbModal
        End If
      End If
    
      'JPD 20030814 Fault 5810
      mblnUserCancelled = mfrmOutput.UserCancelled
    End If
  End If
  
  ' Remove the temporary table.
  If Len(msTempTableName) > 0 Then
    datGeneral.DropUniqueSQLObject msTempTableName, 3
  End If
  
  If Not mfrmOutput Is Nothing Then
    Unload mfrmOutput
  End If
  Set mfrmOutput = Nothing

  Unload frmOutputOptions
  Set frmOutputOptions = Nothing

  Call UtilUpdateLastRun(utlRecordProfile, mlngRecordProfileID)

  'MH20030911 Fault 5811
  fOK = fOK And Not mblnUserCancelled

  If mfNoRecords Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    gobjEventLog.AddDetailEntry msErrorString
    msErrorString = "Completed successfully." & vbCrLf & msErrorString
    fOK = True
  ElseIf fOK And Not mblnUserCancelled Then
    gobjEventLog.ChangeHeaderStatus elsSuccessful
    msErrorString = "Completed successfully."
  ElseIf mblnUserCancelled Then
    gobjEventLog.ChangeHeaderStatus elsCancelled
    msErrorString = "Cancelled by user."
  Else
    'Only details records for failures !
    gobjEventLog.AddDetailEntry msErrorString
    gobjEventLog.ChangeHeaderStatus elsFailed
    msErrorString = "Failed." & vbCrLf & vbCrLf & msErrorString
  End If

  msErrorString = "Record Profile : '" & msRecordProfileName & "' " & msErrorString

  If Not gblnBatchMode Then
    If gobjProgress.Visible Then gobjProgress.CloseProgress
    'If (fOK = False) Or (mfNoRecords = True) Then
    If (fOK = False) Or (mfNoRecords = True) Or Not mblnPreviewOnScreen Then
      COAMsgBox msErrorString, IIf(fOK, vbInformation, vbExclamation) + vbOKOnly, "Record Profile"
    End If
  End If

  Screen.MousePointer = vbDefault
  gobjProgress.ResetBar2
  
  RunRecordProfile = fOK

  Exit Function

RunRecordProfile_ERROR:

  fOK = False
  RunRecordProfile = False
  msErrorString = "Error Whilst Running Record Profile." & vbCrLf & Err.Description
  Resume Next

End Function


Private Function InitialisePreview() As Boolean
  On Error GoTo InitForm_ERROR
  
  Set mfrmOutput = New frmRecordProfilePreview
  
  ' Initialise the progress bar
  With gobjProgress
    If Not gblnBatchMode Then
      .Bar1MaxValue = mrsResults.RecordCount
    Else
      .ResetBar2
      .Bar2MaxValue = (mrsResults.RecordCount + 1)
      .Bar2Caption = "Record Profile : " & msRecordProfileName
    End If
  End With
  
  With mfrmOutput
      
    If mfPrintFilterHeader And (mlngSingleRecordID <= 0) Then
      If (mobjDefinition.BaseTable.FilterID > 0) Then
        .RecordProfileName = msRecordProfileName & " (Base Table filter : " & datGeneral.GetFilterName(mobjDefinition.BaseTable.FilterID) & ")"
      ElseIf (mobjDefinition.BaseTable.PicklistID > 0) Then
        .RecordProfileName = msRecordProfileName & " (Base Table picklist : " & datGeneral.GetPicklistName(mobjDefinition.BaseTable.PicklistID) & ")"
      End If
    Else
      .RecordProfileName = msRecordProfileName
    End If

    .Tag = msRecordProfileName
    
    .Results = mrsResults
    Set .Definition = mobjDefinition
    
    ' Pass the default output options to the calendar preview screen
    .OutputFormat = mlngOutputFormat
    .OutputScreen = mblnOutputScreen
    .OutputPrinter = mblnOutputPrinter
    .OutputPrinterName = mstrOutputPrinterName
    .OutputSave = mblnOutputSave
    .OutputSaveExisting = mlngOutputSaveExisting
    '.OutputSaveFormat = mlngOutputSaveFormat
    .OutputEmail = mblnOutputEmail
    .OutputEmailAddr = mlngOutputEmailAddr
    .OutputEmailSubject = mstrOutputEmailSubject
    .OutputEmailAttachAs = mstrOutputEmailAttachAs
    '.OutputEmailFileFormat = mlngOutputEmailFileFormat
    .OutputFilename = mstrOutputFileName
    
    .IndentRelatedTables = mblnIndentRelatedTables
    .SuppressEmptyRelatedTableTitles = mblnSuppressEmptyRelatedTableTitles
    .SuppressTableRelationshipTitles = mblnSuppressTableRelationshipTitles
  End With
  
  InitialisePreview = True

TidyUpAndExit:
  Exit Function
  
InitForm_ERROR:
  InitialisePreview = False
  msErrorString = "Error whilst initialising Record Profile Preview screen." & vbCrLf & Err.Description
  GoTo TidyUpAndExit
  
End Function

Private Function CheckRecordSet() As Boolean
  On Error GoTo CheckRecordSet_ERROR
  
  If mrsResults.BOF And mrsResults.EOF Then
    CheckRecordSet = False
    msErrorString = "No records meet selection criteria."
    mfNoRecords = True
    Exit Function
  End If
  
  CheckRecordSet = True
  Exit Function
  
CheckRecordSet_ERROR:
  
  msErrorString = "Error while checking returned recordset." & vbCrLf & "(" & Err.Description & ")"
  CheckRecordSet = False
  
End Function


Private Function ExecuteSql() As Boolean
  On Error GoTo ExecuteSQL_ERROR

  Set mrsResults = mclsData.OpenRecordset(msSQL, adOpenStatic, adLockReadOnly)

  ExecuteSql = True
  Exit Function
  
ExecuteSQL_ERROR:
  
  msErrorString = "Error whilst executing SQL statement." & vbCrLf & Err.Description
  ExecuteSql = False
  
End Function


Private Function GenerateFastSQL() As Boolean

  On Error GoTo GenerateFastSQL_ERROR
  
  Dim objTableView As CTablePrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim sRealSource As String
  Dim asViews() As String
  Dim fOK As Boolean
  Dim fColumnOK As Boolean
  Dim fNoSelect As Boolean
  Dim fFound As Boolean
  Dim sColumnCode As String
  Dim sSource As String
  Dim iNextIndex As Integer
  Dim objRecProfTable As clsRecordProfileTabDtl
  Dim objRecProfTable2 As clsRecordProfileTabDtl
  Dim objRecProfRelatedTable As clsRecordProfileTabDtl
  Dim objRecProfColumn As clsRecordProfileColDtl
  Dim avTablesColumns() As Variant
  Dim avTablesJoins() As Variant
  Dim lngLoop As Long
  Dim lngLastTableID As Long
  Dim sFastSQL As String
  Dim iLoop As Integer
  Dim iTableCounter As Integer
  Dim sJoinSQL As String
  Dim sWhereSQL As String
  Dim objRelatedTableView As CTablePrivilege
  Dim objBaseTableView As CTablePrivilege
  Dim avTableViews() As Variant
  Dim rsTemp As New Recordset
  Dim rsTemp2 As New Recordset
  Dim sPickListIDs As String
  Dim sFilterIDs As String
  Dim sRecSelSQL As String
  Dim fColumnRequired As Boolean
  Dim fTableAlreadyInRecProfile As Boolean
  Dim sOrderSQL As String
  Dim sSubOrderSQL As String
  Dim sBaseWhereSQL As String
  Dim sTempTableSQL As String
  Dim iDataType As SQLDataType
  
  ' Set flags with their starting values
  fOK = True
  fNoSelect = False
  ReDim asViews(0)
  Set objBaseTableView = gcoTablePrivileges.FindTableID(mobjDefinition.BaseTable.TableID)

  ' Column 1 = table ID to which this table/view is related
  ' Column 2 = 0 if we're joining a table, 1 if we're joining a view
  ' Column 3 = table/view ID
  ReDim avTableViews(3, 0)

  ' Column 1 = Table ID
  ' Column 2 = Column SELECT code
  ' Column 3 = Column ID
  ' Column 4 = Column name
  ' Column 5 = Column datatype
  ' Column 6 = Column size
  ' Column 7 = Column decimals
  ReDim avTablesColumns(9, 0)
  
  sRecSelSQL = ""
  If mlngSingleRecordID > 0 Then
    sRecSelSQL = CStr(mlngSingleRecordID)
  ElseIf mobjDefinition.BaseTable.PicklistID > 0 Then
    sPickListIDs = ""
    Set rsTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mobjDefinition.BaseTable.PicklistID, adOpenForwardOnly, adLockReadOnly)

    If rsTemp.BOF And rsTemp.EOF Then
      msErrorString = "The selected picklist contains no records."
      GenerateFastSQL = False
      Exit Function
    End If

    Do While Not rsTemp.EOF
      sPickListIDs = sPickListIDs & IIf(Len(sPickListIDs) > 0, ", ", "") & rsTemp.Fields(0)
      rsTemp.MoveNext
    Loop
    rsTemp.Close
    Set rsTemp = Nothing

    sRecSelSQL = sPickListIDs
  ElseIf mobjDefinition.BaseTable.FilterID > 0 Then
    sFilterIDs = ""
    fOK = datGeneral.FilteredIDs(mobjDefinition.BaseTable.FilterID, sFilterIDs)

    ' Generate any UDFs that are used in this filter
    If fOK And gbEnableUDFFunctions Then
      datGeneral.FilterUDFs mobjDefinition.BaseTable.FilterID, mvarUDFsRequired()
    End If

    If fOK Then
      sRecSelSQL = sFilterIDs
    Else
      ' Permission denied on something in the filter.
      msErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mobjDefinition.BaseTable.FilterID) & "' filter."
      GenerateFastSQL = False
      Exit Function
    End If
  End If
  
  ' Loop through the columns collection creating the SELECT and JOIN code
  For Each objRecProfTable In mobjDefinition
    For Each objRecProfColumn In objRecProfTable.Columns
      Set objColumnPrivileges = GetColumnPrivileges(objRecProfTable.TableName)
      sRealSource = gcoTablePrivileges.Item(objRecProfTable.TableName).RealSource
      
      If (objRecProfColumn.ColType = sTYPECODE_COLUMN) Or _
        (objRecProfColumn.ColType = sTYPECODE_ID) Then

        ' Check permission on that column
        fColumnOK = objColumnPrivileges.IsValid(objRecProfColumn.ColumnName)

        If fColumnOK Then
          fColumnOK = objColumnPrivileges.Item(objRecProfColumn.ColumnName).AllowSelect
        End If

        If fColumnOK Then
          ' This column can be read direct from the realSource
          ReDim Preserve avTablesColumns(9, UBound(avTablesColumns, 2) + 1)
          avTablesColumns(1, UBound(avTablesColumns, 2)) = objRecProfTable.TableID
          avTablesColumns(2, UBound(avTablesColumns, 2)) = sRealSource & "." & Trim(objRecProfColumn.ColumnName)
          avTablesColumns(3, UBound(avTablesColumns, 2)) = objRecProfColumn.ID
          avTablesColumns(4, UBound(avTablesColumns, 2)) = objRecProfColumn.ColumnName
'          avTablesColumns(5, UBound(avTablesColumns, 2)) = objRecProfColumn.DataType
'          avTablesColumns(6, UBound(avTablesColumns, 2)) = objRecProfColumn.Size
'          avTablesColumns(7, UBound(avTablesColumns, 2)) = objRecProfColumn.DecPlaces
          Set rsTemp2 = datGeneral.GetColumnDefinition(objRecProfColumn.ID)
          avTablesColumns(5, UBound(avTablesColumns, 2)) = rsTemp2!DataType
          avTablesColumns(6, UBound(avTablesColumns, 2)) = rsTemp2!Size
          avTablesColumns(7, UBound(avTablesColumns, 2)) = rsTemp2!Decimals
          avTablesColumns(8, UBound(avTablesColumns, 2)) = rsTemp2!OLEType
          avTablesColumns(9, UBound(avTablesColumns, 2)) = rsTemp2!Multiline
          rsTemp2.Close
          Set rsTemp2 = Nothing
        Else
          ' This column cannot be read direct. If its from a parent, try parent views
          ' Loop thru the views on the table, seeing if any have read permis for the column

          ReDim asViews(0)
          For Each objTableView In gcoTablePrivileges.Collection
            If (Not objTableView.IsTable) And _
              (objTableView.TableID = objRecProfTable.TableID) And _
              (objTableView.AllowSelect) Then

              sSource = objTableView.ViewName
              sRealSource = gcoTablePrivileges.Item(sSource).RealSource

              ' Get the column permission for the view
              Set objColumnPrivileges = GetColumnPrivileges(sSource)

              ' If we can see the column from this view
              If objColumnPrivileges.IsValid(objRecProfColumn.ColumnName) Then
                If objColumnPrivileges.Item(objRecProfColumn.ColumnName).AllowSelect Then

                  ReDim Preserve asViews(UBound(asViews) + 1)
                  asViews(UBound(asViews)) = objTableView.ViewName

                  ' Check if view has already been added to the array
                  fFound = False
                  For iNextIndex = 1 To UBound(avTableViews, 2)
                    If (avTableViews(1, iNextIndex) = objRecProfTable.TableID) And _
                      (avTableViews(2, iNextIndex) = 1) And _
                      (avTableViews(3, iNextIndex) = objTableView.ViewID) Then

                      fFound = True
                      Exit For
                    End If
                  Next iNextIndex

                  If Not fFound Then
                    ' View has not yet been added, so add it !
                    iNextIndex = UBound(avTableViews, 2) + 1
                    ReDim Preserve avTableViews(3, iNextIndex)
                    avTableViews(1, iNextIndex) = objRecProfTable.TableID
                    avTableViews(2, iNextIndex) = 1
                    avTableViews(3, iNextIndex) = objTableView.ViewID
                  End If
                End If
              End If
            
              Set objColumnPrivileges = Nothing
            End If
          Next objTableView
          Set objTableView = Nothing

          ' Does the user have select permission thru ANY views ?
          If UBound(asViews) = 0 Then
            fNoSelect = True
          Else
            ' Add the column to the column list
            sColumnCode = ""
            For iNextIndex = 1 To UBound(asViews)
              If iNextIndex = 1 Then
                sColumnCode = "CASE"
              End If

              sColumnCode = sColumnCode & _
                " WHEN NOT " & asViews(iNextIndex) & "." & objRecProfColumn.ColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & objRecProfColumn.ColumnName
            Next iNextIndex

            If Len(sColumnCode) > 0 Then
              sColumnCode = sColumnCode & _
                " ELSE NULL" & _
                " END"

              ReDim Preserve avTablesColumns(9, UBound(avTablesColumns, 2) + 1)
              avTablesColumns(1, UBound(avTablesColumns, 2)) = objRecProfTable.TableID
              avTablesColumns(2, UBound(avTablesColumns, 2)) = sColumnCode
              avTablesColumns(3, UBound(avTablesColumns, 2)) = objRecProfColumn.ID
              avTablesColumns(4, UBound(avTablesColumns, 2)) = objRecProfColumn.ColumnName
'              avTablesColumns(5, UBound(avTablesColumns, 2)) = objRecProfColumn.DataType
'              avTablesColumns(6, UBound(avTablesColumns, 2)) = objRecProfColumn.Size
'              avTablesColumns(7, UBound(avTablesColumns, 2)) = objRecProfColumn.DecPlaces
              Set rsTemp2 = datGeneral.GetColumnDefinition(objRecProfColumn.ID)
              avTablesColumns(5, UBound(avTablesColumns, 2)) = rsTemp2!DataType
              avTablesColumns(6, UBound(avTablesColumns, 2)) = rsTemp2!Size
              avTablesColumns(7, UBound(avTablesColumns, 2)) = rsTemp2!Decimals
              avTablesColumns(8, UBound(avTablesColumns, 2)) = rsTemp2!OLEType
              avTablesColumns(9, UBound(avTablesColumns, 2)) = rsTemp2!Multiline
              rsTemp2.Close
              Set rsTemp2 = Nothing
            End If
          End If

          ' If we cant see a column, then get outta here
          If fNoSelect Then
            GenerateFastSQL = False
            msErrorString = vbCrLf & "You do not have permission to see the '" & objRecProfTable.TableName & "." & objRecProfColumn.ColumnName & "' column" & vbCrLf & "either directly or through any views."
            Exit Function
          End If

          If Not fOK Then
            GenerateFastSQL = False
            Exit Function
          End If
        End If
      End If
    
      Set objColumnPrivileges = Nothing
    
    Next objRecProfColumn
    Set objRecProfColumn = Nothing
  Next objRecProfTable
  Set objRecProfTable = Nothing
    
  For Each objRecProfTable In mobjDefinition
    ' Work out the ORDER BY details for this table.
    ' Do not need to do the order by stuff for parents of the base table,
    ' as there will be only on parent record per base table record.
    If objRecProfTable.Relationship <> "PARENT" Then
      sSubOrderSQL = ""
      
      If objRecProfTable.Relationship = "CHILD" Then
        For iLoop = 1 To UBound(avTablesColumns, 2)
          If (avTablesColumns(1, iLoop) = objRecProfTable.TableID) And _
            (avTablesColumns(4, iLoop) = "ID_" & objRecProfTable.RelatedTableID) Then
            sSubOrderSQL = "," & vbCrLf & _
              "   [ASRSysTempCol" & CStr(iLoop) & "]"
            Exit For
          End If
        Next iLoop
      End If
      
      Set rsTemp = datGeneral.GetOrderDefinition(objRecProfTable.OrderID)
    
      Do While Not rsTemp.EOF
        If rsTemp!Type = "O" Then
          ' Do not need to add the order column if it is from a parent of the current table
          ' that is already in the record profile, as the current table's records are reported
          ' per parent table record.
          fColumnRequired = True
        
          If (rsTemp!TableID <> objRecProfTable.TableID) Then
            fTableAlreadyInRecProfile = False
            
            For Each objRecProfTable2 In mobjDefinition
              If objRecProfTable2.TableID = rsTemp!TableID Then
                fTableAlreadyInRecProfile = True
                Exit For
              End If
            
              If objRecProfTable2.TableID = objRecProfTable.TableID Then
                Exit For
              End If
            Next objRecProfTable2
            Set objRecProfTable2 = Nothing
            
            If Not fTableAlreadyInRecProfile Then
              Set objColumnPrivileges = GetColumnPrivileges(rsTemp!TableName)
              sRealSource = gcoTablePrivileges.Item(rsTemp!TableName).RealSource
              
              ' Check permission on that column
              fColumnOK = objColumnPrivileges.IsValid(rsTemp!ColumnName)
              
              If fColumnOK Then
                fColumnOK = objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect
              End If

              Set objColumnPrivileges = Nothing

              If fColumnOK Then
                ' This column can be read direct from the realSource
                ReDim Preserve avTablesColumns(9, UBound(avTablesColumns, 2) + 1)
                avTablesColumns(1, UBound(avTablesColumns, 2)) = objRecProfTable.TableID
                avTablesColumns(2, UBound(avTablesColumns, 2)) = sRealSource & "." & Trim(rsTemp!ColumnName)
                avTablesColumns(3, UBound(avTablesColumns, 2)) = rsTemp!ColumnID
                avTablesColumns(4, UBound(avTablesColumns, 2)) = rsTemp!ColumnName
                
                Set rsTemp2 = datGeneral.GetColumnDefinition(rsTemp!ColumnID)
                avTablesColumns(5, UBound(avTablesColumns, 2)) = rsTemp2!DataType
                avTablesColumns(6, UBound(avTablesColumns, 2)) = rsTemp2!Size
                avTablesColumns(7, UBound(avTablesColumns, 2)) = rsTemp2!Decimals
                avTablesColumns(8, UBound(avTablesColumns, 2)) = rsTemp2!OLEType
                avTablesColumns(9, UBound(avTablesColumns, 2)) = rsTemp2!Multiline
                rsTemp2.Close
                Set rsTemp2 = Nothing

                sSubOrderSQL = sSubOrderSQL & IIf(Len(sSubOrderSQL) = 0, "," & vbCrLf & "   ", ", ") & _
                  "[ASRSysTempCol" & CStr(UBound(avTablesColumns, 2)) & "]" & IIf(rsTemp!Ascending, "", " DESC")
              
                ' Check if table has already been added to the array
                fFound = False
                For iNextIndex = 1 To UBound(avTableViews, 2)
                  If (avTableViews(1, iNextIndex) = objRecProfTable.TableID) And _
                    (avTableViews(2, iNextIndex) = 0) And _
                    (avTableViews(3, iNextIndex) = rsTemp!TableID) Then
              
                    fFound = True
                    Exit For
                  End If
                Next iNextIndex
              
                If Not fFound Then
                  ' Table has not yet been added, so add it !
                  iNextIndex = UBound(avTableViews, 2) + 1
                  ReDim Preserve avTableViews(3, iNextIndex)
                  avTableViews(1, iNextIndex) = objRecProfTable.TableID
                  avTableViews(2, iNextIndex) = 0
                  avTableViews(3, iNextIndex) = rsTemp!TableID
                End If
              Else
                ' This column cannot be read direct. If its from a parent, try parent views
                ' Loop thru the views on the table, seeing if any have read permis for the column
              
                ReDim asViews(0)
                For Each objTableView In gcoTablePrivileges.Collection
                  If (Not objTableView.IsTable) And _
                   (objTableView.TableID = rsTemp!TableID) And _
                    (objTableView.AllowSelect) Then
              
                    sSource = objTableView.ViewName
                    sRealSource = gcoTablePrivileges.Item(sSource).RealSource
              
                    ' Get the column permission for the view
                    Set objColumnPrivileges = GetColumnPrivileges(sSource)
              
                    ' If we can see the column from this view
                    If objColumnPrivileges.IsValid(rsTemp!ColumnName) Then
                      If objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect Then
              
                        ReDim Preserve asViews(UBound(asViews) + 1)
                        asViews(UBound(asViews)) = objTableView.ViewName
              
                        ' Check if view has already been added to the array
                        fFound = False
                        For iNextIndex = 1 To UBound(avTableViews, 2)
                          If (avTableViews(1, iNextIndex) = objRecProfTable.TableID) And _
                            (avTableViews(2, iNextIndex) = 1) And _
                            (avTableViews(3, iNextIndex) = objTableView.ViewID) Then
              
                            fFound = True
                            Exit For
                          End If
                        Next iNextIndex
              
                        If Not fFound Then
                          ' View has not yet been added, so add it !
                          iNextIndex = UBound(avTableViews, 2) + 1
                          ReDim Preserve avTableViews(3, iNextIndex)
                          avTableViews(1, iNextIndex) = objRecProfTable.TableID
                          avTableViews(2, iNextIndex) = 1
                          avTableViews(3, iNextIndex) = objTableView.ViewID
                        End If
                      End If
                    End If
                  
                    Set objColumnPrivileges = Nothing
                  End If
                Next objTableView
                Set objTableView = Nothing
              
                ' Does the user have select permission thru ANY views ?
                If UBound(asViews) > 0 Then
                  ' Add the column to the column list
                  sColumnCode = ""
                  For iNextIndex = 1 To UBound(asViews)
                    If iNextIndex = 1 Then
                      sColumnCode = "CASE"
                    End If
              
                    sColumnCode = sColumnCode & _
                      " WHEN NOT " & asViews(iNextIndex) & "." & rsTemp!ColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & rsTemp!ColumnName
                  Next iNextIndex
              
                  If Len(sColumnCode) > 0 Then
                    sColumnCode = sColumnCode & _
                      " ELSE NULL" & _
                      " END"
              
                    ReDim Preserve avTablesColumns(9, UBound(avTablesColumns, 2) + 1)
                    avTablesColumns(1, UBound(avTablesColumns, 2)) = objRecProfTable.TableID
                    avTablesColumns(2, UBound(avTablesColumns, 2)) = sColumnCode
                    avTablesColumns(3, UBound(avTablesColumns, 2)) = rsTemp!ColumnID
                    avTablesColumns(4, UBound(avTablesColumns, 2)) = rsTemp!ColumnName

                    Set rsTemp2 = datGeneral.GetColumnDefinition(rsTemp!ColumnID)
                    avTablesColumns(5, UBound(avTablesColumns, 2)) = rsTemp2!DataType
                    avTablesColumns(6, UBound(avTablesColumns, 2)) = rsTemp2!Size
                    avTablesColumns(7, UBound(avTablesColumns, 2)) = rsTemp2!Decimals
                    avTablesColumns(8, UBound(avTablesColumns, 2)) = rsTemp2!OLEType
                    avTablesColumns(9, UBound(avTablesColumns, 2)) = rsTemp2!Multiline
                    rsTemp2.Close
                    Set rsTemp2 = Nothing
              
                    sSubOrderSQL = sSubOrderSQL & IIf(Len(sSubOrderSQL) = 0, "," & vbCrLf & "   ", ", ") & _
                      "[ASRSysTempCol" & CStr(UBound(avTablesColumns, 2)) & "]" & IIf(rsTemp!Ascending, "", " DESC")
                  End If
                End If
              End If
            End If
          Else
            fFound = False
            For iLoop = 1 To UBound(avTablesColumns, 2)
              If (avTablesColumns(1, iLoop) = objRecProfTable.TableID) And _
                (avTablesColumns(3, iLoop) = rsTemp!ColumnID) Then
                sSubOrderSQL = sSubOrderSQL & IIf(Len(sSubOrderSQL) = 0, "," & vbCrLf & "   ", ", ") & _
                  "[ASRSysTempCol" & CStr(iLoop) & "]" & IIf(rsTemp!Ascending, "", " DESC")
                fFound = True
                Exit For
              End If
            Next iLoop
          
            If Not fFound Then
              Set objColumnPrivileges = GetColumnPrivileges(objRecProfTable.TableName)
              sRealSource = gcoTablePrivileges.Item(objRecProfTable.TableName).RealSource
              
              ' Check permission on that column
              fColumnOK = objColumnPrivileges.IsValid(rsTemp!ColumnName)
              
              If fColumnOK Then
                fColumnOK = objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect
              End If
              
              Set objColumnPrivileges = Nothing
              
              If fColumnOK Then
                ' This column can be read direct from the realSource
                ReDim Preserve avTablesColumns(9, UBound(avTablesColumns, 2) + 1)
                avTablesColumns(1, UBound(avTablesColumns, 2)) = objRecProfTable.TableID
                avTablesColumns(2, UBound(avTablesColumns, 2)) = sRealSource & "." & Trim(rsTemp!ColumnName)
                avTablesColumns(3, UBound(avTablesColumns, 2)) = rsTemp!ColumnID
                avTablesColumns(4, UBound(avTablesColumns, 2)) = rsTemp!ColumnName

                Set rsTemp2 = datGeneral.GetColumnDefinition(rsTemp!ColumnID)
                avTablesColumns(5, UBound(avTablesColumns, 2)) = rsTemp2!DataType
                avTablesColumns(6, UBound(avTablesColumns, 2)) = rsTemp2!Size
                avTablesColumns(7, UBound(avTablesColumns, 2)) = rsTemp2!Decimals
                avTablesColumns(8, UBound(avTablesColumns, 2)) = rsTemp2!OLEType
                avTablesColumns(9, UBound(avTablesColumns, 2)) = rsTemp2!Multiline
                rsTemp2.Close
                Set rsTemp2 = Nothing
              
                sSubOrderSQL = sSubOrderSQL & IIf(Len(sSubOrderSQL) = 0, "," & vbCrLf & "   ", ", ") & _
                  "[ASRSysTempCol" & CStr(UBound(avTablesColumns, 2)) & "]" & IIf(rsTemp!Ascending, "", " DESC")
              Else
                ' This column cannot be read direct. If its from a parent, try parent views
                ' Loop thru the views on the table, seeing if any have read permis for the column
              
                ReDim asViews(0)
                For Each objTableView In gcoTablePrivileges.Collection
                  If (Not objTableView.IsTable) And _
                   (objTableView.TableID = objRecProfTable.TableID) And _
                    (objTableView.AllowSelect) Then
              
                    sSource = objTableView.ViewName
                    sRealSource = gcoTablePrivileges.Item(sSource).RealSource
              
                    ' Get the column permission for the view
                    Set objColumnPrivileges = GetColumnPrivileges(sSource)
              
                    ' If we can see the column from this view
                    If objColumnPrivileges.IsValid(rsTemp!ColumnName) Then
                      If objColumnPrivileges.Item(rsTemp!ColumnName).AllowSelect Then
              
                        ReDim Preserve asViews(UBound(asViews) + 1)
                        asViews(UBound(asViews)) = objTableView.ViewName
              
                        ' Check if view has already been added to the array
                        fFound = False
                        For iNextIndex = 1 To UBound(avTableViews, 2)
                          If (avTableViews(1, iNextIndex) = objRecProfTable.TableID) And _
                            (avTableViews(2, iNextIndex) = 1) And _
                            (avTableViews(3, iNextIndex) = objTableView.ViewID) Then
              
                            fFound = True
                            Exit For
                          End If
                        Next iNextIndex
              
                        If Not fFound Then
                          ' View has not yet been added, so add it !
                          iNextIndex = UBound(avTableViews, 2) + 1
                          ReDim Preserve avTableViews(3, iNextIndex)
                          avTableViews(1, iNextIndex) = objRecProfTable.TableID
                          avTableViews(2, iNextIndex) = 1
                          avTableViews(3, iNextIndex) = objTableView.ViewID
                        End If
                      End If
                    End If
                  
                    Set objColumnPrivileges = Nothing
                  End If
                Next objTableView
                Set objTableView = Nothing
              
                ' Does the user have select permission thru ANY views ?
                If UBound(asViews) > 0 Then
                  ' Add the column to the column list
                  sColumnCode = ""
                  For iNextIndex = 1 To UBound(asViews)
                    If iNextIndex = 1 Then
                      sColumnCode = "CASE"
                    End If
              
                    sColumnCode = sColumnCode & _
                      " WHEN NOT " & asViews(iNextIndex) & "." & rsTemp!ColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & rsTemp!ColumnName
                  Next iNextIndex
              
                  If Len(sColumnCode) > 0 Then
                    sColumnCode = sColumnCode & _
                      " ELSE NULL" & _
                      " END"
              
                    ReDim Preserve avTablesColumns(9, UBound(avTablesColumns, 2) + 1)
                    avTablesColumns(1, UBound(avTablesColumns, 2)) = objRecProfTable.TableID
                    avTablesColumns(2, UBound(avTablesColumns, 2)) = sColumnCode
                    avTablesColumns(3, UBound(avTablesColumns, 2)) = rsTemp!ColumnID
                    avTablesColumns(4, UBound(avTablesColumns, 2)) = rsTemp!ColumnName

                    Set rsTemp2 = datGeneral.GetColumnDefinition(rsTemp!ColumnID)
                    avTablesColumns(5, UBound(avTablesColumns, 2)) = rsTemp2!DataType
                    avTablesColumns(6, UBound(avTablesColumns, 2)) = rsTemp2!Size
                    avTablesColumns(7, UBound(avTablesColumns, 2)) = rsTemp2!Decimals
                    avTablesColumns(8, UBound(avTablesColumns, 2)) = rsTemp2!OLEType
                    avTablesColumns(9, UBound(avTablesColumns, 2)) = rsTemp2!Multiline
                    rsTemp2.Close
                    Set rsTemp2 = Nothing
                  
                    sSubOrderSQL = sSubOrderSQL & IIf(Len(sSubOrderSQL) = 0, "," & vbCrLf & "   ", ", ") & _
                      "[ASRSysTempCol" & CStr(UBound(avTablesColumns, 2)) & "]" & IIf(rsTemp!Ascending, "", " DESC")
                  End If
                End If
              End If
            End If
          End If
        End If
        
        rsTemp.MoveNext
      Loop
    
      rsTemp.Close
      Set rsTemp = Nothing
      
      ' Always add the ID column last.
      If objRecProfTable.Relationship = "CHILD" Then
        For iLoop = 1 To UBound(avTablesColumns, 2)
          If (avTablesColumns(1, iLoop) = objRecProfTable.TableID) And _
            (avTablesColumns(4, iLoop) = "ID") Then
            sSubOrderSQL = sSubOrderSQL & ", [ASRSysTempCol" & CStr(iLoop) & "]"
            Exit For
          End If
        Next iLoop
      End If
      
      sOrderSQL = sOrderSQL & sSubOrderSQL
      
    End If
  Next objRecProfTable
  Set objRecProfTable = Nothing
    
  '
  ' Construct the SQL Select statement.
  '
  sFastSQL = ""
  iTableCounter = 0
  sBaseWhereSQL = ""
  
  msTempTableName = datGeneral.UniqueSQLObjectName("ASRSysTempRecordProfile", 3)
  'MH20060301 Fault 10845
  'sTempTableSQL = "CREATE TABLE [" & datGeneral.UserNameForSQL & "].[" & msTempTableName & "] (" & _
    "[ASRSysTableOrder] [int]"
  sTempTableSQL = "CREATE TABLE [" & gsUserName & "].[" & msTempTableName & "] (" & _
    "[ASRSysTableOrder] [int]"
    
  For lngLoop = 1 To UBound(avTablesColumns, 2)
    sTempTableSQL = sTempTableSQL & _
      ", [ASRSysTempCol" & CStr(lngLoop) & "] "
      
    iDataType = CInt(avTablesColumns(5, lngLoop))
    Select Case iDataType
      Case sqlVarChar
        If avTablesColumns(9, lngLoop) = True Then
          sTempTableSQL = sTempTableSQL & "[varchar] (MAX)"
        Else
          sTempTableSQL = sTempTableSQL & "[varchar] (" & CStr(avTablesColumns(6, lngLoop)) & ")"
        End If
      Case sqlBoolean
        sTempTableSQL = sTempTableSQL & "[bit]"
      Case sqlDate
        sTempTableSQL = sTempTableSQL & "[datetime]"
      Case sqlInteger
        sTempTableSQL = sTempTableSQL & "[int]"
      Case sqlNumeric
        sTempTableSQL = sTempTableSQL & "[numeric](" & CStr(avTablesColumns(6, lngLoop)) & "," & CStr(avTablesColumns(7, lngLoop)) & ")"
      Case sqlOle, sqlVarBinary
        If avTablesColumns(8, lngLoop) = OLE_EMBEDDED Then
          sTempTableSQL = sTempTableSQL & "[image]"
        Else
          sTempTableSQL = sTempTableSQL & "[varchar] (255)"
        End If
      Case sqlLongVarChar
        sTempTableSQL = sTempTableSQL & "[varchar] (14)"
    End Select
    
    sTempTableSQL = sTempTableSQL & " NULL"
      
  Next lngLoop
  sTempTableSQL = sTempTableSQL & ")"
  datGeneral.GetRecords sTempTableSQL

  For Each objRecProfTable In mobjDefinition
    ' Loop through the tables in the Record Profile, constructing the
    ' SQL SELECT statement for each one, UNIONing each tables code together.
    
    ' Initialise the SELECT statement (UNION it with existing SELECT code if required).
    'MH20060301 Fault 10845
    'sFastSQL = "INSERT INTO [" & datGeneral.UserNameForSQL & "].[" & msTempTableName & "] SELECT "
    sFastSQL = "INSERT INTO [" & gsUserName & "].[" & msTempTableName & "] SELECT "
    
    ' SELECT a dummy column to represent the table's order in the RecordProfile.
    sFastSQL = sFastSQL & CStr(iTableCounter) & " AS [ASRSysTableOrder]"

    ' Add the column SELECT code.
    lngLastTableID = 0
    For lngLoop = 1 To UBound(avTablesColumns, 2)
      sFastSQL = sFastSQL & ","
      
      ' Bit poncey this bit. Add a carriage-return so that all columns for a particular table
      ' are grouped on one line.
      If lngLastTableID <> avTablesColumns(1, lngLoop) Then
        sFastSQL = sFastSQL & vbCrLf
        lngLastTableID = avTablesColumns(1, lngLoop)
      End If
      
      ' Add the column SELECT code if the column is in the current table; else just add 'null'.
      If (objRecProfTable.TableID = CLng(avTablesColumns(1, lngLoop))) Or _
        ((mobjDefinition.BaseTable.TableID = CLng(avTablesColumns(1, lngLoop))) And (avTablesColumns(4, lngLoop) = "ID")) Then
        sFastSQL = sFastSQL & " " & _
          CStr(avTablesColumns(2, lngLoop))
      Else
        sFastSQL = sFastSQL & " " & _
          "null"
      End If
      
      sFastSQL = sFastSQL & " " & _
        " AS [ASRSysTempCol" & CStr(lngLoop) & "]"
    Next lngLoop
    
    ' Add the FROM code.
    sFastSQL = sFastSQL & vbCrLf & _
      " FROM " & gcoTablePrivileges.Item(mobjDefinition.BaseTable.TableName).RealSource
      
    ' Join on the required tables.
    sJoinSQL = ""
    
    ' Join related tables.
    Set objRecProfRelatedTable = objRecProfTable
    Do While objRecProfRelatedTable.TableID <> mobjDefinition.BaseTable.TableID
      Set objTableView = gcoTablePrivileges.FindTableID(objRecProfRelatedTable.TableID)
      Set objRelatedTableView = gcoTablePrivileges.FindTableID(objRecProfRelatedTable.RelatedTableID)
      
      If objRecProfRelatedTable.Relationship = "CHILD" Then
        sJoinSQL = " INNER JOIN " & objTableView.RealSource & _
          " ON " & objRelatedTableView.RealSource & ".ID = " & objTableView.RealSource & ".ID_" & objRelatedTableView.TableID & IIf(Len(sJoinSQL) > 0, vbCrLf, "") & _
          sJoinSQL
      Else
        sJoinSQL = " INNER JOIN " & objTableView.RealSource & _
          " ON " & objRelatedTableView.RealSource & ".ID_" & objTableView.TableID & " = " & objTableView.RealSource & ".ID" & IIf(Len(sJoinSQL) > 0, vbCrLf, "") & _
          sJoinSQL
      End If
      
      Set objRecProfRelatedTable = mobjDefinition.Item(CStr(objRecProfRelatedTable.RelatedTableID))
    
      Set objTableView = Nothing
      Set objRelatedTableView = Nothing
    Loop
    Set objRecProfRelatedTable = Nothing
    
    ' Join views.
    sWhereSQL = ""
    If objRecProfTable.TableID <> mobjDefinition.BaseTable.TableID Then
      sWhereSQL = sBaseWhereSQL
    End If
    
    For iLoop = 1 To UBound(avTableViews, 2)
      If objRecProfTable.TableID = CLng(avTableViews(1, iLoop)) Then
        If CInt(avTableViews(2, iLoop)) = 1 Then
          Set objTableView = gcoTablePrivileges.FindViewID(CLng(avTableViews(3, iLoop)))
        Else
          Set objTableView = gcoTablePrivileges.FindTableID(CLng(avTableViews(3, iLoop)))
        End If
        Set objRelatedTableView = gcoTablePrivileges.FindTableID(objRecProfTable.TableID)
        
        If objRecProfTable.TableID = objTableView.TableID Then
          sJoinSQL = sJoinSQL & IIf(Len(sJoinSQL) > 0, vbCrLf, "") & _
            " LEFT OUTER JOIN " & objTableView.RealSource & _
            " ON " & objRelatedTableView.RealSource & ".ID = " & objTableView.RealSource & ".ID"
        
          If Not objRelatedTableView.AllowSelect Then
            sWhereSQL = sWhereSQL & IIf(Len(sWhereSQL) > 0, vbCrLf & "   OR ", "") & _
              "(" & objRelatedTableView.RealSource & ".ID IN (SELECT " & objTableView.RealSource & ".ID FROM " & objTableView.RealSource & "))"
          End If
        Else
          sJoinSQL = sJoinSQL & IIf(Len(sJoinSQL) > 0, vbCrLf, "") & _
            " LEFT OUTER JOIN " & objTableView.RealSource & _
            " ON " & objRelatedTableView.RealSource & ".ID_" & CStr(objTableView.TableID) & " = " & objTableView.RealSource & ".ID"
        End If
      
        Set objTableView = Nothing
        Set objRelatedTableView = Nothing
      End If
    Next iLoop
    
    If Len(sWhereSQL) > 0 Then
      sWhereSQL = "(" & sWhereSQL & ")"
    End If
    
    ' Add on the JOIN code.
    If Len(sJoinSQL) > 0 Then
      sFastSQL = sFastSQL & vbCrLf & _
        sJoinSQL
    End If
    
    If objRecProfTable.TableID = mobjDefinition.BaseTable.TableID Then
      If (mlngSingleRecordID > 0) Or _
        (mobjDefinition.BaseTable.PicklistID > 0) Or _
        (mobjDefinition.BaseTable.FilterID > 0) Then
        
        sWhereSQL = sWhereSQL & IIf(Len(sWhereSQL) > 0, vbCrLf & "   AND ", "") & _
          objBaseTableView.RealSource & ".ID IN (" & sRecSelSQL & ")"
      End If
    End If
    
    If (objRecProfTable.TableID <> mobjDefinition.BaseTable.TableID) And _
      (objRecProfTable.FilterID > 0) Then
      
      sFilterIDs = ""
      Set objTableView = gcoTablePrivileges.FindTableID(objRecProfTable.TableID)
      fOK = datGeneral.FilteredIDs(objRecProfTable.FilterID, sFilterIDs)

      If fOK Then
        sWhereSQL = sWhereSQL & IIf(Len(sWhereSQL) > 0, vbCrLf & "   AND ", "") & _
          objTableView.RealSource & ".ID IN (" & sFilterIDs & ")"
      Else
        ' Permission denied on something in the filter.
        msErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(objRecProfTable.FilterID) & "' filter."
        GenerateFastSQL = False
        Exit Function
      End If
    
      Set objTableView = Nothing
    End If
    
    If objRecProfTable.TableID = mobjDefinition.BaseTable.TableID Then
      sBaseWhereSQL = sWhereSQL
    End If
    
    ' Add on the WHERE code.
    If Len(sWhereSQL) > 0 Then
      sFastSQL = sFastSQL & vbCrLf & _
        " WHERE (" & sWhereSQL & ")"
    End If
    
    UDFFunctions mvarUDFsRequired, True
    datGeneral.ExecuteSql sFastSQL, msErrorString
    UDFFunctions mvarUDFsRequired, False
    
    If msErrorString <> vbNullString Then
      GenerateFastSQL = False
      ' Remove the temporary table.
      datGeneral.DropUniqueSQLObject msTempTableName, 3
      Exit Function
    End If

    iTableCounter = iTableCounter + 1
  Next objRecProfTable
  Set objRecProfTable = Nothing
  
  Set objBaseTableView = Nothing
  
  'MH20060301 Fault 10845
  'sFastSQL = "SELECT * FROM [" & datGeneral.UserNameForSQL & "].[" & msTempTableName & "]"
  sFastSQL = "SELECT * FROM [" & gsUserName & "].[" & msTempTableName & "]"

  ' Add on the ORDER BY code.
  sFastSQL = sFastSQL & vbCrLf & vbCrLf & _
    " ORDER BY [ASRSysTableOrder]" & sOrderSQL

  msSQL = sFastSQL
  
  GenerateFastSQL = True
    
  Exit Function
  
GenerateFastSQL_ERROR:
  
  GenerateFastSQL = False
  msErrorString = "Error whilst generating the SQL query." & vbCrLf & Err.Description
  
End Function


Private Function GetRecordProfileColumns() As Boolean
  ' Purpose : This function loads record profile details into
  '           arrays and leaves the details recordset reference there

  On Error GoTo GetRecordProfileColumns_ERROR
  
  Dim strTempSQL As String
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim sTempTableName As String
  Dim lngTempTableID As Long
  Dim objColumn As CColumnPrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim alngRelatingColumns() As Long
  Dim rsDetails As New Recordset
  Dim objTable As clsRecordProfileTabDtl
  Dim objRecProfColumn As clsRecordProfileColDtl
  Dim iColumnCounter As Integer
  Dim lngTempColumnID As Long
  Dim iTempDataType As SQLDataType
  Dim sTempColumnName As String
  'Dim sTempDigitSeparator As String
  'Dim iTempDigitGrouping As Integer
  Dim bTempThousandSeparator As Boolean
  Dim bBlankIfZero As Boolean
  Dim iOLEType As DataMgr.OLEType
  Dim fIsSystemColumn As Boolean
  
  ReDim alngRelatingColumns(2, 0)
  ' column 1 = parent table ID
  ' column 2 = position of child's ID_ column in the recordset
  
  strTempSQL = "SELECT ASRSysRecordProfileDetails.*," & _
    " ASRSysColumns.columnName, ASRSysColumns.dataType, ASRSysColumns.Use1000Separator, ASRSysColumns.BlankIfZero,  ASRSysColumns.OLEType " & _
    " FROM ASRSysRecordProfileDetails" & _
    " LEFT OUTER JOIN ASRSysColumns" & _
    "   ON (ASRSysRecordProfileDetails.type = '" & sTYPECODE_COLUMN & "' AND ASRSysRecordProfileDetails.columnID = ASRSysColumns.columnID)" & _
    " WHERE ASRSysRecordProfileDetails.recordProfileID = " & mlngRecordProfileID & _
    " ORDER BY ASRSysRecordProfileDetails.tableID, ASRSysRecordProfileDetails.sequence"

  Set rsDetails = mclsData.OpenRecordset(strTempSQL, adOpenKeyset, adLockReadOnly)
  With rsDetails
    If .BOF And .EOF Then
      GetRecordProfileColumns = False
      msErrorString = "No columns found in the specified record profile definition." & vbCrLf & "Please remove this definition and create a new one."
      Exit Function
    End If

    Do Until .EOF
      Set objTable = mobjDefinition.Item(CStr(!TableID))
      
      If (!Type = sTYPECODE_COLUMN) Then
        lngTempColumnID = !ColumnID
        iTempDataType = !DataType
        sTempColumnName = !ColumnName
        'sTempDigitSeparator = ","
        'iTempDigitGrouping = IIf(IsNull(!Use1000Separator), 0, 3)
        bTempThousandSeparator = IIf(IsNull(!Use1000Separator), 0, !Use1000Separator)
        bBlankIfZero = IIf(IsNull(!BlankIfZero), 0, !BlankIfZero)
        iOLEType = IIf(IsNull(!OLEType), OLE_LOCAL, !OLEType)
      Else
        lngTempColumnID = 0
        iTempDataType = 0
        sTempColumnName = ""
        'sTempDigitSeparator = ""
        'iTempDigitGrouping = 0
        bTempThousandSeparator = False
        iOLEType = OLE_LOCAL
      End If

      Set objRecProfColumn = objTable.Columns.Add(CStr(objTable.Columns.Count + 1), _
        !Type, _
        lngTempColumnID, _
        !Heading, _
        !Size, _
        !dp, _
        iTempDataType, _
        !TableID, _
        sTempColumnName, _
        objTable.Columns.Count + 1)
        
      With objRecProfColumn
        .Displayed = True
        .PositionInRecordset = 0
        .TableName = IIf(rsDetails!Type = sTYPECODE_COLUMN, datGeneral.GetTableName(CLng(rsDetails!TableID)), 0)
        .ThousandSeparator = bTempThousandSeparator
        .BlankIfZero = bBlankIfZero
        .OLEType = iOLEType
        '.DigitSeparator = sTempDigitSeparator
        '.DigitGrouping = iTempDigitGrouping
      End With
      
      Set objRecProfColumn = Nothing

      If !Type = sTYPECODE_HEADING Then
        objTable.HasHeadings = True
      End If
      
      If iTempDataType = sqlVarBinary Then
        objTable.PhotoCount = objTable.PhotoCount + 1
      End If
      
      .MoveNext
    
      Set objTable = Nothing
    Loop
    
    .MoveFirst
  End With

  rsDetails.Close
  Set rsDetails = Nothing
  
  ' Remove training headings for horizontal grid displays as they look funny.
  For Each objTable In mobjDefinition
    If objTable.Orientation = giHORIZONTAL Then
      Do While objTable.Columns.Item(objTable.Columns.Count).ColType = sTYPECODE_HEADING
        objTable.Columns.Remove (CStr(objTable.Columns.Count))
      Loop
    End If
  Next objTable
  Set objTable = Nothing
  
  '******************************************************************************
  ' Add the ID columns for the tables
  '******************************************************************************
  
  ' Add the tables ID and ID_<n> columns
  For Each objTable In mobjDefinition
    sTempTableName = objTable.TableName
    lngTempTableID = objTable.TableID
    
    Set objColumnPrivileges = GetColumnPrivileges(objTable.TableName)
    For Each objColumn In objColumnPrivileges
      If (objColumn.ColumnType = colSystem) Then
        objTable.Columns.Add CStr(objTable.Columns.Count + 1), _
          sTYPECODE_ID, _
          objColumn.ColumnID, _
          "", _
          0, _
          0, _
          sqlInteger, _
          objTable.TableID, _
          objColumn.ColumnName, _
          objTable.Columns.Count + 1
      End If
    Next objColumn
    Set objColumn = Nothing
    Set objColumnPrivileges = Nothing
  Next objTable
  Set objTable = Nothing
  
  iColumnCounter = 1
  
  For Each objTable In mobjDefinition
    For Each objRecProfColumn In objTable.Columns
      If (objRecProfColumn.ColType = sTYPECODE_COLUMN) Or _
        (objRecProfColumn.ColType = sTYPECODE_ID) Then
      
        objRecProfColumn.PositionInRecordset = iColumnCounter
        
        If objRecProfColumn.ColumnName = "ID" Then
          objTable.IDPosition = iColumnCounter
        ElseIf UCase(objRecProfColumn.ColumnName) = UCase("ID_" & Trim(Str(objTable.RelatedTableID))) Then
          objTable.RelatedTableIDPosition = iColumnCounter
        ElseIf (UCase(Left(objRecProfColumn.ColumnName, 3)) = "ID_") Then
          'JPD 20060224 Fault 10839
          fIsSystemColumn = False
          strTempSQL = "SELECT columnType FROM ASRSysColumns WHERE columnid = " & CStr(objRecProfColumn.ID)
          Set rsDetails = mclsData.OpenRecordset(strTempSQL, adOpenForwardOnly, adLockReadOnly)

          If Not rsDetails.BOF And Not rsDetails.EOF Then
            fIsSystemColumn = (rsDetails(0) = colSystem)
          End If
        
          rsDetails.Close
          Set rsDetails = Nothing
        
          If fIsSystemColumn Then
            ReDim Preserve alngRelatingColumns(2, UBound(alngRelatingColumns, 2) + 1)
            alngRelatingColumns(1, UBound(alngRelatingColumns, 2)) = CLng(Mid(objRecProfColumn.ColumnName, 4))
            alngRelatingColumns(2, UBound(alngRelatingColumns, 2)) = iColumnCounter
          End If
        End If
      
        iColumnCounter = iColumnCounter + 1
      End If
    Next objRecProfColumn
    Set objRecProfColumn = Nothing
  Next objTable
  Set objTable = Nothing
    
  For Each objTable In mobjDefinition
    If objTable.Relationship = "PARENT" Then
      For iLoop2 = 0 To UBound(alngRelatingColumns, 2)
        If objTable.TableID = alngRelatingColumns(1, iLoop2) Then
          objTable.RelatedTableIDPosition = alngRelatingColumns(2, iLoop2)
          Exit For
        End If
      Next iLoop2
    End If
  Next objTable
  Set objTable = Nothing
  
  GetRecordProfileColumns = True
  Exit Function

GetRecordProfileColumns_ERROR:

  GetRecordProfileColumns = False
  msErrorString = "Error whilst retrieving the details recordsets'." & vbCrLf & Err.Description
  
End Function


Private Function GetDataType(lColumnID As Long) As Long

  'Needed to be created as the one in datgeneral needs on tableid
  
  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "Select DataType From ASRSysColumns Where ColumnID = " & lColumnID
  Set rsTemp = New Recordset
  rsTemp.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  GetDataType = rsTemp(0)
  
  rsTemp.Close
  Set rsTemp = Nothing

End Function

Private Function GetRecordProfileTables() As Boolean
  
  ' Purpose : This function retrieves the basic definition details
  '           and stores it in module level variables
  
  On Error GoTo GetRecordProfileTables_ERROR

  Dim rsTemp_Definition As Recordset
  Dim strSQL As String
  Dim lngTempOrderID As Long
  Dim objTempTable As clsRecordProfileTabDtl
  Dim objTempTable2 As clsRecordProfileTabDtl
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim lngFilterID As Long
  Dim lblnReportPackMode As Boolean
  
  lblnReportPackMode = gblnReportPackMode
  
  strSQL = "SELECT ASRSysRecordProfileName.*, ASRSysTables.recordDescExprID" & _
    " FROM ASRSysRecordProfileName" & _
    " INNER JOIN ASRSysTables ON ASRSysRecordProfileName.baseTable = ASRSysTables.tableID" & _
    " WHERE recordProfileID = " & mlngRecordProfileID & " "
  Set rsTemp_Definition = mclsData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  With rsTemp_Definition
    If .BOF And .EOF Then
      GetRecordProfileTables = False
      msErrorString = "Could not find specified Record Profile definition !"
      .Close
      Set rsTemp_Definition = Nothing
      Exit Function
    End If

    msRecordProfileName = !Name
    mstrRecordProfileDescription = !Description
    mbDefinitionOwner = (LCase(Trim(gsUserName)) = LCase(Trim(!UserName)))
  
    'New Default Output Variables
'    mlngOutputFormat = IIf(IsNull(!OutputFormat), 0, !OutputFormat)
'    mblnOutputScreen = IIf(IsNull(!OutputScreen), False, !OutputScreen)
'    mblnOutputPrinter = IIf(IsNull(!OutputPrinter), False, !OutputPrinter)
'    mstrOutputPrinterName = IIf(IsNull(!OutputPrinterName), "", !OutputPrinterName)
'    mblnOutputSave = IIf(IsNull(!OutputSave), False, !OutputSave)
'    mlngOutputSaveExisting = IIf(IsNull(!OutputSaveExisting), 0, !OutputSaveExisting)
'    'mlngOutputSaveFormat = IIf(IsNull(!OutputSaveFormat), 0, !OutputSaveFormat)
'    mblnOutputEmail = IIf(IsNull(!OutputEmail), False, !OutputEmail)
'    mlngOutputEmailAddr = IIf(IsNull(!OutputEmailAddr), 0, !OutputEmailAddr)
'    mstrOutputEmailSubject = IIf(IsNull(!OutputEmailSubject), "", !OutputEmailSubject)
'    mstrOutputEmailAttachAs = IIf(IsNull(!OutputEmailAttachAs), "", !OutputEmailAttachAs)
'    'mlngOutputEmailFileFormat = IIf(IsNull(!OutputEmailFileFormat), 0, !OutputEmailFileFormat)
'    mstrOutputFileName = IIf(IsNull(!OutputFilename), "", !OutputFilename)
  
    'Change Output Options to Report Pack owning these Jobs if in Report Pack mode
    mblnPreviewOnScreen = IIf(lblnReportPackMode, False, !OutputPreview Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))
    mblnOutputScreen = IIf(lblnReportPackMode, False, (IIf(IsNull(!OutputScreen), False, !OutputScreen)))
    mlngOutputFormat = IIf(lblnReportPackMode, mlngOutputFormat, (IIf(IsNull(!OutputFormat), 0, !OutputFormat)))
    mblnOutputPrinter = IIf(lblnReportPackMode, mblnOutputPrinter, IIf(IsNull(!OutputPrinter), False, !OutputPrinter))
    mstrOutputPrinterName = IIf(lblnReportPackMode, mstrOutputPrinterName, IIf(IsNull(!OutputPrinterName), "", !OutputPrinterName))
    mblnOutputSave = IIf(lblnReportPackMode, mblnOutputSave, IIf(IsNull(!OutputSave), False, !OutputSave))
    mlngOutputSaveExisting = IIf(lblnReportPackMode, mlngOutputSaveExisting, IIf(IsNull(!OutputSaveExisting), 0, !OutputSaveExisting))
    mblnOutputEmail = IIf(lblnReportPackMode, mblnOutputEmail, IIf(IsNull(!OutputEmail), False, !OutputEmail))
    mlngOutputEmailAddr = IIf(lblnReportPackMode, mlngOutputEmailAddr, IIf(IsNull(!OutputEmailAddr), 0, !OutputEmailAddr))
    mstrOutputEmailSubject = IIf(lblnReportPackMode, mstrOutputEmailSubject, IIf(IsNull(!OutputEmailSubject), "", !OutputEmailSubject))
    mstrOutputEmailAttachAs = IIf(lblnReportPackMode, mstrOutputEmailAttachAs, IIf(IsNull(!OutputEmailAttachAs), "", !OutputEmailAttachAs))
    mstrOutputFileName = IIf(lblnReportPackMode, mstrOutputFileName, IIf(IsNull(!OutputFilename), "", !OutputFilename))
    mlngOverrideFilterID = IIf(lblnReportPackMode, mlngOverrideFilterID, 0)
    
    
    
    mblnIndentRelatedTables = IIf(IsNull(!IndentRelatedTables), False, !IndentRelatedTables)
    mblnSuppressEmptyRelatedTableTitles = IIf(IsNull(!SuppressEmptyRelatedTableTitles), False, !SuppressEmptyRelatedTableTitles)
    mblnSuppressTableRelationshipTitles = IIf(IsNull(!SuppressTableRelationshipTitles), False, !SuppressTableRelationshipTitles)
    mfPrintFilterHeader = IIf(IsNull(!PrintFilterHeader), False, !PrintFilterHeader)
    
    lngTempOrderID = IIf(IsNull(!OrderID), 0, !OrderID)
    
    If lngTempOrderID = 0 Then lngTempOrderID = datGeneral.GetDefaultOrder(!BaseTable)
    ' Use Override filter if there is one
    If gblnReportPackMode Then
      lngFilterID = IIf(mlngOverrideFilterID <> 0, mlngOverrideFilterID, IIf(IsNull(!FilterID), 0, !FilterID))
    Else
      lngFilterID = !FilterID
    End If
    
    Set objTempTable = mobjDefinition.Add(!BaseTable, _
      datGeneral.GetTableName(!BaseTable), _
      IIf(IsNull(!RecordDescExprID), 0, !RecordDescExprID), _
      lngFilterID, _
      IIf(IsNull(!PicklistID), 0, !PicklistID), _
      lngTempOrderID, _
      0, _
      !Orientation, _
      !PageBreak, _
      0)
      
    objTempTable.Relationship = "BASE"
    
    Set objTempTable = Nothing

    .Close
  End With
  
  Set rsTemp_Definition = Nothing

  strSQL = "SELECT ASRSysRecordProfileTables.tableID," & _
    " ASRSysRecordProfileTables.FilterID," & _
    " ASRSysRecordProfileTables.OrderID," & _
    " ASRSysRecordProfileTables.MaxRecords," & _
    " ASRSysRecordProfileTables.orientation," & _
    " ASRSysRecordProfileTables.pageBreak," & _
    " ASRSysRecordProfileTables.sequence," & _
    " ASRSysTables.TableName," & _
    " ASRSysTables.recordDescExprID" & _
    " FROM ASRSysRecordProfileTables" & _
    " INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysRecordProfileTables.tableID" & _
    " WHERE ASRSysRecordProfileTables.recordProfileID = " & mlngRecordProfileID & _
    " ORDER BY ASRSysRecordProfileTables.sequence"

  Set rsTemp_Definition = mclsData.OpenRecordset(strSQL, adOpenStatic, adLockReadOnly)

  With rsTemp_Definition
    Do Until .EOF
      lngTempOrderID = IIf(IsNull(!OrderID), 0, !OrderID)
      If lngTempOrderID = 0 Then
        lngTempOrderID = datGeneral.GetDefaultOrder(!TableID)
      End If
      
      mobjDefinition.Add !TableID, _
        !TableName, _
        IIf(IsNull(!RecordDescExprID), 0, !RecordDescExprID), _
        lngFilterID, _
        0, _
        lngTempOrderID, _
        !MaxRecords, _
        !Orientation, _
        !PageBreak, _
        !Sequence
      .MoveNext
    Loop
    
    .Close
  End With
  Set rsTemp_Definition = Nothing

  If Not IsRecordSelectionValid Then
    GetRecordProfileTables = False
    Exit Function
  End If

  ' Determine the order the tables will be reported in.
  miIndex = 1
  GetRecordProfileTableOrder "PARENT", mobjDefinition.BaseTable.TableID, 1
  GetRecordProfileTableOrder "CHILD", mobjDefinition.BaseTable.TableID, 1
  
  ' Order the collection
  For iLoop = 2 To mobjDefinition.Count
    For iLoop2 = 2 To mobjDefinition.Count
      Set objTempTable = mobjDefinition.ItemByPosition(iLoop2)
      
      If objTempTable.Sequence = (iLoop - 1) Then
        mobjDefinition.Remove CStr(objTempTable.TableID)
        
        Set objTempTable2 = mobjDefinition.Add(objTempTable.TableID, _
          objTempTable.TableName, _
          objTempTable.RecordDescriptionID, _
          objTempTable.FilterID, _
          0, _
          objTempTable.OrderID, _
          objTempTable.MaxRecords, _
          objTempTable.Orientation, _
          objTempTable.PageBreak, _
          objTempTable.Sequence)
        
        With objTempTable2
          .Generation = objTempTable.Generation
          .HasChildren = objTempTable.HasChildren
          .LastRecordID = 0
          .RelatedTableID = objTempTable.RelatedTableID
          .RelatedTableIDPosition = 0
          .Relationship = objTempTable.Relationship
        End With
      
        Set objTempTable2 = Nothing
      End If
    
      Set objTempTable = Nothing
    Next iLoop2
  Next iLoop
  
  GetRecordProfileTables = True
  
TidyAndExit:
  
  Set rsTemp_Definition = Nothing

Exit Function

GetRecordProfileTables_ERROR:

  GetRecordProfileTables = False
  msErrorString = "Error whilst reading the Record Profile definition !" & vbCrLf & Err.Description
  Resume TidyAndExit

End Function


Public Property Get PreviewOnScreen() As Boolean
  PreviewOnScreen = ((msErrorString = vbNullString And mblnPreviewOnScreen) And Not gblnBatchMode)
End Property

Private Function IsRecordSelectionValid() As Boolean
  ' Check all of the record selection parameters are still valid.
  Dim iLoop As Integer
  Dim lngFilterID As Long
  Dim lngPicklistID As Long
  Dim iResult As RecordSelectionValidityCodes
  
  ' Validate the base table filter.
  lngFilterID = mobjDefinition.BaseTable.FilterID
  iResult = ValidateRecordSelection(REC_SEL_FILTER, lngFilterID)
  Select Case iResult
    Case REC_SEL_VALID_DELETED
      msErrorString = "The base table filter used in this definition has been deleted by another user."
    Case REC_SEL_VALID_INVALID
      msErrorString = "The base table filter used in this definition is invalid."
    Case REC_SEL_VALID_HIDDENBYOTHER
      If Not gfCurrentUserIsSysSecMgr Then
        msErrorString = "The base table filter used in this definition has been made hidden by another user."
      End If
  End Select
  
  ' Validate the base table picklist.
  If Len(msErrorString) = 0 Then
    lngPicklistID = mobjDefinition.BaseTable.PicklistID
    iResult = ValidateRecordSelection(REC_SEL_PICKLIST, lngPicklistID)
    Select Case iResult
      Case REC_SEL_VALID_DELETED
        msErrorString = "The base table picklist used in this definition has been deleted by another user."
      Case REC_SEL_VALID_INVALID
        msErrorString = "The base table picklist used in this definition is invalid."
      Case REC_SEL_VALID_HIDDENBYOTHER
        If Not gfCurrentUserIsSysSecMgr Then
          msErrorString = "The base table picklist used in this definition has been made hidden by another user."
        End If
      End Select
  End If
  
  ' Validate the related tables filters.
  If Len(msErrorString) = 0 Then
    For iLoop = 2 To mobjDefinition.Count
      lngFilterID = mobjDefinition.ItemByPosition(iLoop).FilterID
  
      iResult = ValidateRecordSelection(REC_SEL_FILTER, lngFilterID)
      Select Case iResult
        Case REC_SEL_VALID_DELETED
          msErrorString = "The '" & mobjDefinition.ItemByPosition(iLoop).TableName & "' table filter used in this definition has been deleted by another user."
        Case REC_SEL_VALID_INVALID
          msErrorString = "The '" & mobjDefinition.ItemByPosition(iLoop).TableName & "' table filter used in this definition is invalid."
        Case REC_SEL_VALID_HIDDENBYOTHER
          If Not gfCurrentUserIsSysSecMgr Then
            msErrorString = "The '" & mobjDefinition.ItemByPosition(iLoop).TableName & "' table filter used in this definition has been made hidden by another user."
          End If
      End Select
      
      If Len(msErrorString) > 0 Then
        Exit For
      End If
    Next iLoop
  End If

  IsRecordSelectionValid = (Len(msErrorString) = 0)
  InvalidPicklistFilter = Not IsRecordSelectionValid
  
End Function

    Public Sub SetOutputParameters( _
          lngOutputFormat As Long, _
          blnOutputScreen As Boolean, _
          blnOutputPrinter As Boolean, _
          strOutputPrinterName As String, _
          blnOutputSave As Boolean, _
          lngOutputSaveExisting As Long, _
          blnOutputEmail As Boolean, _
          lngOutputEmailAddr As Long, _
          strOutputEmailSubject As String, _
          strOutputEmailAttachAs As String, _
          strOutputFilename As String, _
          blnPreviewOnScreen As Boolean, _
          blnChkPicklistFilter As Boolean, _
          Optional strOutputTitlePage As String, _
          Optional strOutputReportPackTitle As String, _
          Optional strOutputOverrideFilter As String, _
          Optional blnOutputTOC As Boolean, _
          Optional blnOutputCoverSheet As Boolean, _
          Optional lngOverrideFilterID As Long)

  mlngOutputFormat = lngOutputFormat
  mblnOutputScreen = blnOutputScreen
  mblnOutputPrinter = blnOutputPrinter
  mstrOutputPrinterName = strOutputPrinterName
  mblnOutputSave = blnOutputSave
  mlngOutputSaveExisting = lngOutputSaveExisting
  mblnOutputEmail = blnOutputEmail
  mlngOutputEmailAddr = lngOutputEmailAddr
  mstrOutputEmailSubject = strOutputEmailSubject
  mstrOutputEmailAttachAs = strOutputEmailAttachAs
  mstrOutputFileName = strOutputFilename
  mblnChkPicklistFilter = blnChkPicklistFilter
  mblnPreviewOnScreen = (blnPreviewOnScreen Or (mlngOutputFormat = fmtDataOnly And mblnOutputScreen))
  mstrOutputTitlePage = IIf(IsMissing(strOutputTitlePage), giEXPRVALUE_CHARACTER, strOutputTitlePage)
  mstrOutputReportPackTitle = IIf(IsMissing(strOutputReportPackTitle), giEXPRVALUE_CHARACTER, strOutputReportPackTitle)
  mstrOutputOverrideFilter = IIf(IsMissing(strOutputOverrideFilter), giEXPRVALUE_CHARACTER, strOutputOverrideFilter)
  mblnOutputTOC = IIf(IsMissing(blnOutputTOC), giEXPRVALUE_CHARACTER, blnOutputTOC)
  mblnOutputCoverSheet = IIf(IsMissing(blnOutputCoverSheet), giEXPRVALUE_CHARACTER, blnOutputCoverSheet)
  mlngOverrideFilterID = IIf(IsMissing(lngOverrideFilterID), giEXPRVALUE_CHARACTER, lngOverrideFilterID)
End Sub

Public Property Get InvalidPicklistFilter() As Boolean
  InvalidPicklistFilter = mblnInvalidPicklistFilter
End Property
Private Property Let InvalidPicklistFilter(bValid As Boolean)
  mblnInvalidPicklistFilter = bValid
End Property

Public Property Get ErrorString() As String
  ErrorString = msErrorString
End Property

Public Property Let ErrorString(ByVal sErrorString As String)
  msErrorString = sErrorString
End Property


Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property


Public Property Let RecordProfileID(ByVal lNewID As Long)
  mlngRecordProfileID = lNewID
  
End Property


Public Property Get RecordProfileID() As Long
  RecordProfileID = mlngRecordProfileID
  
End Property

Private Sub Class_Initialize()
  Set mclsData = New DataMgr.clsDataAccess
  Set mclsGeneral = New DataMgr.clsGeneral
  ReDim mavTableViews(4, 0)
  
  Set mobjDefinition = New clsRecordProfileTabDtls

End Sub


Private Sub Class_Terminate()
  Set mclsData = Nothing
  Set mclsGeneral = Nothing

  Set mobjDefinition = Nothing

End Sub



