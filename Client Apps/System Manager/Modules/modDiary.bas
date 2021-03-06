Attribute VB_Name = "modDiary"
Option Explicit

Private rsDiaryLinks As ADODB.Recordset


Public Function OpenDiaryRecordsets()

  Set rsDiaryLinks = New ADODB.Recordset
  rsDiaryLinks.Open "ASRSysDiaryLinks", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

End Function

  
Public Function SaveDiaryLinksForColumn(lngColumnID As Long)

  ' Add the diary link values.
  With recDiaryEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst

      Do While Not .EOF
        If !ColumnID = lngColumnID Then
          rsDiaryLinks.AddNew
          rsDiaryLinks!diaryID = !diaryID
          rsDiaryLinks!ColumnID = !ColumnID
          rsDiaryLinks!Comment = !Comment
          rsDiaryLinks!Offset = !Offset
          rsDiaryLinks!Period = !Period
          rsDiaryLinks!Reminder = !Reminder
          rsDiaryLinks!FilterID = IIf(IsNull(!FilterID), 0, !FilterID)
          rsDiaryLinks!EffectiveDate = IIf(IsNull(!EffectiveDate), "01/01/1980", !EffectiveDate)
          rsDiaryLinks!CheckLeavingDate = IIf(IsNull(!CheckLeavingDate), True, !CheckLeavingDate)
          rsDiaryLinks.Update
        End If

        .MoveNext
      Loop
    End If
  End With

End Function


Public Function CloseDiaryRecordsets()

  If rsDiaryLinks.State = adStateClosed Then
    rsDiaryLinks.Close
  End If
  Set rsDiaryLinks = Nothing

End Function



Public Function GetSQLForRecordDescription(lngRecordDescExprID As Long) As String

  'This is used within the triggers (audit section)
  'and within the diary stored procedures

  GetSQLForRecordDescription = _
    "    /* ---------------------- */" & vbNewLine & _
    "    /* Get Record Description */" & vbNewLine & _
    "    /* ---------------------- */" & vbNewLine & _
    "    DECLARE @recordDesc char(255)," & vbNewLine & _
    "            @oldValue varchar(255)," & vbNewLine & _
    "            @newValue varchar(255)" & vbNewLine & vbNewLine & _
    "    /* Evaluate the inserted record's description (if it is defined). */" & vbNewLine & _
    "    IF EXISTS (SELECT *" & vbNewLine & _
    "        FROM sysobjects" & vbNewLine & _
    "        WHERE type = 'P'" & vbNewLine & _
    "        AND name = 'sp_ASRExpr_" & Trim(Str(lngRecordDescExprID)) & "')" & vbNewLine & _
    "    BEGIN" & vbNewLine & _
    "        EXEC @hResult = sp_ASRExpr_" & Trim(Str(lngRecordDescExprID)) & " @recordDesc OUTPUT, @recordID" & vbNewLine & _
    "        IF @hResult <> 0 SET @recordDesc = ''" & vbNewLine & _
    "        SET @recordDesc = CONVERT(varchar(255), @recordDesc)" & vbNewLine & _
    "    END" & vbNewLine & _
    "    ELSE SET @recordDesc = ''" & vbNewLine & vbNewLine
    
End Function





Public Function CreateDiaryProcsForTable(pLngCurrentTableID As Long, sCurrentTable As String, lngRecordDescExprID As Long) As Integer
  ' JPD 10/5/00 - Changed from being a subroutine to being a function that returns the number of
  ' diary linked columns in the given table.
  Dim iDiaryLinkedColumns As Integer
  Dim recDiaryLinks As New ADODB.Recordset
  Dim sDiaryProcedureSQL As String
  Dim sInsertDiaryCode As String
  Dim sDiaryPeriod As String
  Dim sSQL As String
  Dim strDiaryProcName As String
  'Dim strDiaryRebuildName As String
  Dim strLinkFilter As String
  Dim lngColumnID As Long
  Dim lngLinkID As Long

  Dim strSQLLeavingDate As String
  Dim blnLeavingDate As Boolean

  
  strSQLLeavingDate = GetSQLForLeavingDate(pLngCurrentTableID, sCurrentTable)

  
  
  ' JPD 10/5/00 - initialise the number of diary linked columns in the given table.
  iDiaryLinkedColumns = 0

  strDiaryProcName = "dbo.spASRDiary_" & CStr(pLngCurrentTableID)
  'strDiaryRebuildName = "dbo.sp_ASRDiaryRebuild_" & CStr(pLngCurrentTableID)
  
  
  'Loop through all the columns on this table
  sInsertDiaryCode = vbNullString
  With recColEdit
    .Index = "idxName"
    .Seek ">=", pLngCurrentTableID
      
    If Not .NoMatch Then
        
      Do While Not .EOF
          
        If !TableID <> pLngCurrentTableID Then
          Exit Do
        End If
            
        If recColEdit!DataType = dtTIMESTAMP Then
          
          sSQL = "SELECT * FROM ASRSysDiaryLinks " & _
                 "WHERE ColumnID = '" & recColEdit!ColumnID & "'"
          recDiaryLinks.Open sSQL, gADOCon, adOpenForwardOnly

          lngColumnID = recColEdit!ColumnID

          'Loop through all of the diary links for this column
          'and create an insert command for each
          With recDiaryLinks
            ' JPD 10/5/00 - increment the number of diary linked columns in the given table.
            If Not (.EOF And .BOF) Then
              iDiaryLinkedColumns = iDiaryLinkedColumns + 1
            End If
            
            '.MoveFirst

            If Not .EOF Then
              sInsertDiaryCode = sInsertDiaryCode & _
                "    /* " & recColEdit!ColumnName & " triggers */" & vbNewLine & _
                "    SELECT @oldDateValue = " & recColEdit!ColumnName & " FROM " & sCurrentTable & vbNewLine & _
                "    WHERE @recordid = ID" & vbNewLine & vbNewLine
              
              
              Do While Not .EOF

                sInsertDiaryCode = sInsertDiaryCode & _
                  "    SET @Done = 0" & vbNewLine & _
                  "    IF (NOT @oldDateValue IS NULL)" & vbNewLine & _
                  "    BEGIN" & vbNewLine & vbNewLine

                'Select Case .rdoColumns("Period").Value
                'Case 0: sDiaryPeriod = "day"
                'Case 1: sDiaryPeriod = "week"
                'Case 2: sDiaryPeriod = "month"
                'Case 3: sDiaryPeriod = "year"
                'End Select
                lngLinkID = IIf(IsNull(.Fields("diaryID")), 0, .Fields("diaryID"))
                sDiaryPeriod = TimePeriod(.Fields("Period").value)

                sInsertDiaryCode = sInsertDiaryCode & _
                  "      SELECT @NewDateValue = DATEADD(" & sDiaryPeriod & ", " & .Fields("Offset").value & ", @oldDateValue)" & vbNewLine & _
                  "      SELECT @DiaryComment = CONVERT(varchar(255), RTRIM(@recordDesc) + ': " & Replace(.Fields("Comment").value, "'", "''") & "')" & vbNewLine

                sInsertDiaryCode = sInsertDiaryCode & _
                  "      IF DateDiff(day, '" & Replace(Format(.Fields("EffectiveDate"), "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "', @newDateValue) >= 0 AND DateDiff(day, '12/31/2999', @newDateValue) <= 0" & vbNewLine & _
                  "      BEGIN" & vbNewLine
                
                'sInsertDiaryCode = sInsertDiaryCode & _
                  "      IF DateDiff(day, '" & Format(.rdoColumns("EffectiveDate"), "mm/dd/yyyy") & "', @newDateValue) >= 0" & vbnewline & _
                  "      BEGIN" & vbnewline

                  '"      IF DateDiff(day, @oldestDateAllowed, @newDateValue) >= 0" & vbnewline & _

                sInsertDiaryCode = sInsertDiaryCode & _
                  "        IF (DateDiff(day, @purgeDate, @newDateValue) >= 0) OR @purgeDate IS NULL" & vbNewLine & _
                  "        BEGIN" & vbNewLine


                If strSQLLeavingDate <> vbNullString And Abs(.Fields("CheckLeavingDate")) > 0 Then
                  blnLeavingDate = True
                  sInsertDiaryCode = sInsertDiaryCode & _
                    "          IF (DateDiff(day, @EmployeeLeavingDate, @newDateValue) <= 0) OR @EmployeeLeavingDate IS NULL" & vbNewLine & _
                    "          BEGIN" & vbNewLine
                End If


                If .Fields("FilterID") > 0 Then
                  strLinkFilter = GetSQLFilter(.Fields("FilterID"), sCurrentTable)

                  'MH20010711
                  'recColEdit.Index = "idxColumnID"
                  'recColEdit.Seek "=", lngColumnID
                  recColEdit.Index = "idxName"
                  recColEdit.Seek ">=", pLngCurrentTableID
                  Do While recColEdit!ColumnID <> lngColumnID And Not recColEdit.EOF
                    recColEdit.MoveNext
                  Loop
                  
                  
                  If strLinkFilter <> vbNullString Then
                    sInsertDiaryCode = sInsertDiaryCode & _
                      "            IF " & strLinkFilter & vbNewLine & _
                      "            BEGIN" & vbNewLine
                  End If
                Else
                  strLinkFilter = vbNullString

                End If

                'Only alarm if required and event is in the future
                If Abs(.Fields("Reminder")) > 0 Then
                  sInsertDiaryCode = sInsertDiaryCode & _
                    "              /* If prior to today don't alarm */" & vbNewLine & _
                    "              SELECT @Alarm = CASE WHEN " & _
                                   "(datediff(day,getdate(),@newDateValue) >= 0) " & _
                                   "THEN 1 ELSE 0 END" & vbNewLine & vbNewLine
                Else
                  sInsertDiaryCode = sInsertDiaryCode & _
                    "              /* This event is never alarmed */" & vbNewLine & _
                    "              SELECT @Alarm = 0" & vbNewLine & vbNewLine
                End If
                
                
                sInsertDiaryCode = sInsertDiaryCode & _
                  "             SET @Done = 1" & vbNewLine & vbNewLine

                sInsertDiaryCode = sInsertDiaryCode & _
                  "             IF EXISTS(SELECT * FROM ASRSysDiaryEvents" & vbNewLine & _
                  "             WHERE LinkID = " & CStr(lngLinkID) & " AND RowID = @recordID)" & vbNewLine
                
                sInsertDiaryCode = sInsertDiaryCode & _
                  "              UPDATE ASRSysDiaryEvents SET " & vbNewLine & _
                  "                EventTitle = @DiaryComment," & vbNewLine & _
                  "                EventDate = @NewDateValue," & vbNewLine & _
                  "                ColumnValue = @oldDateValue," & vbNewLine


                'MH20060210 Fault 10651
                If Abs(.Fields("Reminder")) > 0 Then
                  sInsertDiaryCode = sInsertDiaryCode & _
                    "                Alarm = CASE WHEN @Alarm = 1 THEN 1 ELSE Alarm END" & vbNewLine
                Else
                  sInsertDiaryCode = sInsertDiaryCode & _
                    "                Alarm = 0" & vbNewLine
                End If


                sInsertDiaryCode = sInsertDiaryCode & _
                  "              WHERE LinkID = " & CStr(lngLinkID) & vbNewLine & _
                  "                AND RowID = @recordID" & vbNewLine
                
                sInsertDiaryCode = sInsertDiaryCode & _
                  "             ELSE" & vbNewLine

                sInsertDiaryCode = sInsertDiaryCode & _
                  "              INSERT INTO ASRSysDiaryEvents" & vbNewLine & _
                  "                (LinkID, TableID, ColumnID, RowID, EventTitle, EventDate, " & _
                                   "ColumnValue, Alarm, UserName, Access)" & vbNewLine

                sInsertDiaryCode = sInsertDiaryCode & _
                  "              VALUES" & vbNewLine & _
                  "               (" & CStr(lngLinkID) & ", " & _
                                  CStr(pLngCurrentTableID) & ", " & _
                                  CStr(lngColumnID) & ", " & _
                                  "@recordID, " & _
                                  "@DiaryComment, " & _
                                  "@NewDateValue, " & _
                                  "@oldDateValue, " & _
                                  "@Alarm, " & _
                                  "'System', '" & ACCESS_READONLY & "')" & vbNewLine & vbNewLine


                If strLinkFilter <> vbNullString Then
                  sInsertDiaryCode = sInsertDiaryCode & _
                    "          END" & vbNewLine
                End If

                If strSQLLeavingDate <> vbNullString And Abs(.Fields("CheckLeavingDate")) > 0 Then
                  sInsertDiaryCode = sInsertDiaryCode & _
                    "        END" & vbNewLine
                End If

                sInsertDiaryCode = sInsertDiaryCode & _
                  "        END" & vbNewLine & _
                  "      END" & vbNewLine & _
                  "    END" & vbNewLine

                sInsertDiaryCode = sInsertDiaryCode & _
                  "    IF @Done = 0" & vbNewLine & _
                  "      DELETE FROM ASRSysDiaryEvents" & vbNewLine & _
                  "      WHERE LinkID = " & CStr(lngLinkID) & " AND RowID = @recordID" & vbNewLine & vbNewLine

                .MoveNext   'Next Diary Link
              Loop

            End If
              
          End With
          recDiaryLinks.Close

        End If
        
        .MoveNext   'Next column
      Loop
          
    End If
  End With


  ' Drop procedure
  DropProcedure strDiaryProcName


  If sInsertDiaryCode <> vbNullString Then

    sSQL = "/* ---------------------------------------------------------------- */" & vbNewLine _
        & "/* Diary module stored procedure.          */" & vbNewLine _
        & "/* Automatically generated by the System manager.   */" & vbNewLine _
        & "/* ---------------------------------------------------------------- */" & vbNewLine _
        & "CREATE PROCEDURE " & strDiaryProcName _
        & "(@recordID int)" & vbNewLine _
        & "AS" & vbNewLine _
        & "BEGIN" & vbNewLine & vbNewLine _
        & "  DECLARE @hResult int," & vbNewLine _
        & "          @DiaryComment varchar(255)," & vbNewLine _
        & "          @oldDateValue datetime," & vbNewLine _
        & "          @NewDateValue datetime," & vbNewLine _
        & "          @Alarm int," & vbNewLine _
        & "          @purgeDate datetime," & vbNewLine _
        & "          @Done bit" & vbNewLine & vbNewLine

    sSQL = sSQL & _
      "  EXEC [dbo].[sp_ASRPurgeDate] @purgedate OUTPUT, 'DIARYSYS'" & vbNewLine & vbNewLine & _
      GetSQLForRecordDescription(lngRecordDescExprID) & vbNewLine & vbNewLine

    If blnLeavingDate Then
      sSQL = sSQL & _
        "  DECLARE @EmployeeLeavingDate datetime" & vbNewLine & _
        strSQLLeavingDate
    End If

    sSQL = sSQL & _
      sInsertDiaryCode & _
      "END"

    gADOCon.Execute sSQL, , adExecuteNoRecords
    
    End If
  
  ' JPD 10/5/00 - return the number of diary linked columns in the given table.
  CreateDiaryProcsForTable = iDiaryLinkedColumns
  
  Set recDiaryLinks = Nothing
  
End Function


Private Function GetSQLFilter(lngFilterID As Long, sCurrentTable As String) As String

  Dim fOK As Boolean
  Dim objExpr As CExpression
  Dim strFilterRunTimeCode As String

  GetSQLFilter = vbNullString

  'Filter
  Set objExpr = New CExpression
  With objExpr

    objExpr.ExpressionID = lngFilterID
    objExpr.ConstructExpression
    fOK = objExpr.RuntimeFilterCode(strFilterRunTimeCode, False)

    strFilterRunTimeCode = Replace(strFilterRunTimeCode, vbNewLine, " ")
      
    GetSQLFilter = "@recordID IN " & _
          "(" & strFilterRunTimeCode & ")"
    
  End With
  Set objExpr = Nothing

End Function


Private Function GetSQLForLeavingDate(lngCurrentTable As Long, strCurrentTable As String) As String

  Dim lngPersonnelTableID As Long
  'Dim lngStartDateID As Long
  Dim lngLeavingDateID As Long
  Dim strSQL As String
  Dim blnChildOfPers As Boolean
  
  
  ' Check if Leaving Date column in module setup
  With recModuleSetup
    .Index = "idxModuleParameter"
    
    ' Get the Personnel table ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    If .NoMatch Then
      lngPersonnelTableID = 0
    Else
      lngPersonnelTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    If lngPersonnelTableID = 0 Then
      Exit Function
    End If


    '' Get the Start Date column ID.
    '.Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_STARTDATE
    'If .NoMatch Then
    '  lngStartDateID = 0
    'Else
    '  lngStartDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    'End If

    ' Get the Leaving Date column ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_LEAVINGDATE
    If .NoMatch Then
      lngLeavingDateID = 0
    Else
      lngLeavingDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

  End With


  'If not personnel table then check if child of personnel
  If lngCurrentTable <> lngPersonnelTableID Then

    With recTabEdit
      .Index = "idxName"
      
      If Not (.BOF And .EOF) Then
        .MoveFirst
      End If
      
      Do While Not .EOF()
        If (Not !Deleted) And _
          (!TableID <> lngPersonnelTableID) And _
          (!TableType = iTabChild) Then

          recRelEdit.Index = "idxParentID"
          recRelEdit.Seek "=", lngPersonnelTableID, lngCurrentTable
          blnChildOfPers = (Not recRelEdit.NoMatch)

        End If
        
        .MoveNext
      Loop

    End With

  End If

  If lngCurrentTable = lngPersonnelTableID Or blnChildOfPers Then

      With recColEdit
        .Index = "idxColumnID"
        .Seek "=", lngLeavingDateID
    
        If Not .NoMatch Then
          With recTabEdit
            .Index = "idxTableID"
            .Seek "=", recColEdit!TableID
          
            If Not .NoMatch Then
              GetSQLForLeavingDate = "  SELECT @EmployeeLeavingDate = [" & Trim(recColEdit!ColumnName) & "]" & _
                       " FROM [" & Trim(recTabEdit!TableName) & "]" & _
                       " WHERE ID = "
              
              If blnChildOfPers Then
                GetSQLForLeavingDate = GetSQLForLeavingDate & _
                    "(SELECT [ID_" & CStr(lngPersonnelTableID) & "]" & _
                    " FROM [" & strCurrentTable & "] WHERE ID = @recordid)" & vbNewLine & vbNewLine
              Else
                GetSQLForLeavingDate = GetSQLForLeavingDate & "@recordid" & vbNewLine & vbNewLine
              End If

            End If
          End With
        
        End If
      End With

  End If

End Function


Public Function TableHasDiaryLinks(lngTableID As Long) As Boolean
  
  Dim blnResult As Boolean
  Dim lngColumnID As Long

  blnResult = False

  With recColEdit
    .Index = "idxName"
    .Seek ">=", lngTableID
      
    If Not .NoMatch Then
      Do While Not .EOF And blnResult = False
        If !TableID <> lngTableID Then
          Exit Do
        End If
            
        blnResult = ColumnHasDiaryLinks(recColEdit!ColumnID)
      
        .MoveNext
      Loop
    End If

  End With

  TableHasDiaryLinks = blnResult

End Function


Public Function ColumnHasDiaryLinks(lngColumnID As Long) As Boolean

  Dim objDiaryLink As cDiaryLink
  Dim blnResult As Boolean

  blnResult = False

  With recDiaryEdit
    .Index = "idxColumnID"
    .Seek "=", lngColumnID
    
    If Not .NoMatch Then
      Do While Not .EOF And blnResult = False
        If !ColumnID <> lngColumnID Then
          Exit Do
        End If
        
        Set objDiaryLink = New cDiaryLink
        objDiaryLink.DiaryLinkId = !diaryID
        blnResult = objDiaryLink.ReadDiaryLink
        Set objDiaryLink = Nothing
        
        .MoveNext
      Loop
    End If
  End With

  ColumnHasDiaryLinks = blnResult

End Function


Public Function DiaryRebuild() As Boolean
  
  Dim rsIDs As ADODB.Recordset
  Dim strSQL As String

  Dim varArray() As Variant
  Dim strCaption As String
  Dim lngIndex As Long


  With recTabEdit
    .Index = "idxName"

    If Not (.BOF And .EOF) Then
      .MoveFirst

      strSQL = vbNullString
      Do While Not .EOF()
        If (Not !Deleted) Then
          If TableHasDiaryLinks(!TableID) Then
            If strSQL <> vbNullString Then
              strSQL = strSQL & "UNION" & vbNewLine
            End If
            strSQL = strSQL & "SELECT ID, " & CStr(!TableID) & " as 'TableID', '" & !TableName & "' as 'TableName' FROM [" & !TableName & "]" & vbNewLine
          End If
        End If
        
        .MoveNext
      Loop

    End If
  
  End With

  If strSQL = vbNullString Then
    Exit Function
  End If
  strSQL = strSQL & " ORDER BY 'TableID'"


  'Get all of the IDs and read into an array
  Set rsIDs = New ADODB.Recordset
  rsIDs.Open strSQL, gADOCon, adOpenStatic, adLockReadOnly, adCmdText
  If rsIDs.BOF And rsIDs.EOF Then
    rsIDs.Close
    Set rsIDs = Nothing
  Else
    varArray = rsIDs.GetRows()
    rsIDs.Close
    Set rsIDs = Nothing

    gobjProgress.ResetBar2
    gobjProgress.Bar2MaxValue = UBound(varArray, 2) + 1
    For lngIndex = 0 To UBound(varArray, 2)

      If strCaption <> CStr(varArray(2, lngIndex)) Then
        strCaption = CStr(varArray(2, lngIndex))
        OutputCurrentProcess2 strCaption
      End If
      strSQL = "EXEC spASRDiary_" & CStr(varArray(1, lngIndex)) & " " & CStr(varArray(0, lngIndex))
      gADOCon.Execute strSQL

      gobjProgress.UpdateProgress2

    Next

  End If

End Function











