Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Linq
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Public Class clsDataTransferRun
    Inherits BaseForDMI

    Private fOK As Boolean
    Private mblnUserCancelled As Boolean
    Private rsDataTransferColumns As DataTable
    Private mstrStatusMessage As String
    Private mlngSelectedID As Integer
    Private mlngTableViews(,) As Integer
    Private mstrProcedureName As String
    Private mstrGetMaxID As String

    Private mlngSuccessCount As Integer
    Private mlngFailCount As Integer

    Private mstrSQLInsert As String
    Private mstrSQLSelect As String
    Private mstrSQLFrom As String
    Private mstrSQLJoin As String
    Private mstrSQLWhere As String
    Private mstrSQLTable() As String

    Private mstrTransferName As String
    Private mlngFromTableID As Integer
    Private mstrFromTableName As String
    Private mlngRecordDescExprID As Integer
    Private mlngToTableID As Integer
    Private mstrToTableName As String

    Private mlngFilterID As Integer
    Private mlngPicklistID As Integer

    Private mbDefinitionOwner As Boolean
    Private mbLoggingDTSuccess As Boolean

    ' Array holding the User Defined functions that are needed for this report
    Private mastrUDFsRequired() As String

    Private mstrPicklistFilterIDs As String

    Private Function IsRecordSelectionValid() As Boolean

        Dim iResult As RecordSelectionValidityCodes

        ' Filter
        If mlngFilterID > 0 Then
            iResult = ValidateRecordSelection(RecordSelectionType.Filter, mlngFilterID)
            Select Case iResult
                Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                    mstrStatusMessage = "The base table filter used in this definition has been deleted by another user."
                Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                    mstrStatusMessage = "The base table filter used in this definition is invalid."
                Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                    If Not _login.IsSystemOrSecurityAdmin Then
                        mstrStatusMessage = "The base table filter used in this definition has been made hidden by another user."
                    End If
            End Select
        ElseIf mlngPicklistID > 0 Then
            iResult = ValidateRecordSelection(RecordSelectionType.Picklist, mlngPicklistID)
            Select Case iResult
                Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                    mstrStatusMessage = "The base table picklist used in this definition has been deleted by another user."
                Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                    mstrStatusMessage = "The base table picklist used in this definition is invalid."
                Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                    If Not _login.IsSystemOrSecurityAdmin Then
                        mstrStatusMessage = "The base table picklist used in this definition has been made hidden by another user."
                    End If
            End Select
        End If

        IsRecordSelectionValid = (Len(mstrStatusMessage) = 0)

    End Function

    Public ReadOnly Property StatusMessage As String
        Get
            Return mstrStatusMessage
        End Get
    End Property

    Private Function Records(lngRec As Integer) As String
        Return CStr(lngRec) & IIf(lngRec <> 1, " records", " record").ToString
    End Function

    Public Function ExecuteDataTransfer(lngSelectedID As Integer, Optional strRecordIDs As String = "") As Boolean

        ReDim mastrUDFsRequired(0)

        mlngSelectedID = lngSelectedID
        mstrPicklistFilterIDs = strRecordIDs

        fOK = True

        If fOK Then Call GetDataTransferDetails()

        mbLoggingDTSuccess = CBool(GetUserSetting("LogEvents", "Data_Transfer_Success", False))

        CheckIfTransferIsValid()

        Logs.AddHeader(EventLog_Type.eltDataTransfer, mstrTransferName)

        CreateInsertTableArray()
        mstrProcedureName = General.UniqueSQLObjectName("sp_ASRTempDataTransfer", 4)
        If fOK Then CreateStoredProcedure()
        If fOK Then ProcessRecords()
        If fOK Then General.UDFFunctions(mastrUDFsRequired, True)

        TidyUpAndExit()
        OutputJobStatus()

        Return fOK

    End Function

    Private Sub OutputJobStatus()

        Dim blnNoRecords As Boolean

        AccessLog.UtilUpdateLastRun(UtilityType.utlDataTransfer, mlngSelectedID)

        If mlngFailCount > 0 And Not mblnUserCancelled Then
            fOK = False
        End If

        If fOK Then
            Logs.ChangeHeaderStatus(EventLog_Status.elsSuccessful, mlngSuccessCount, mlngFailCount)
            mstrStatusMessage = "Completed successfully."

            blnNoRecords = (mlngSuccessCount = 0 And mlngFailCount = 0)

            If blnNoRecords Then
                mstrStatusMessage = mstrStatusMessage & vbNewLine & "No records meet selection criteria."
                Logs.AddDetailEntry(mstrStatusMessage)
                fOK = True
            End If

        ElseIf mblnUserCancelled Then
            Logs.ChangeHeaderStatus(EventLog_Status.elsCancelled, mlngSuccessCount, mlngFailCount)
            mstrStatusMessage = "Cancelled by user."

        Else
            Logs.ChangeHeaderStatus(EventLog_Status.elsFailed, mlngSuccessCount, mlngFailCount)
            If mstrStatusMessage <> vbNullString Then
                'Only details records for failures with description !
                Logs.AddDetailEntry(mstrStatusMessage)
                mstrStatusMessage = "Failed." & vbNewLine & vbNewLine & mstrStatusMessage
            Else
                mstrStatusMessage = "Failed." & vbNewLine
            End If

        End If


        If Not blnNoRecords Then
            mstrStatusMessage = mstrStatusMessage & vbNewLine & vbNewLine & Records(mlngSuccessCount) & " successfully transferred."

            'If mlngFailCount > 0 And Not mlngSingleRecordID > 0 Then
            If mlngFailCount > 0 Then
                'If mstrPicklistFilterIDs = "" Or InStr(mstrPicklistFilterIDs, ",") > 0 Then
                mstrStatusMessage = mstrStatusMessage & vbNewLine & vbNewLine & Records(mlngFailCount) & " failed during transfer."
                'End If
            End If
        End If

        mstrStatusMessage = "Data Transfer : '" & mstrTransferName & "' " & mstrStatusMessage

    End Sub

    Private Sub GetDataTransferDetails()

        Dim rsTemp As DataTable
        Dim strSQL As String

        Try


            strSQL = "SELECT ASRSysDatatransferName.*, " & "FromTable.TableName AS FromTableName, " & "ToTable.TableName AS ToTableName, " & "FromTable.RecordDescExprID AS RecDescExprID " & "FROM ASRSysDatatransferName " & "JOIN ASRSysTables FromTable " & "  ON FromTable.TableID = ASRSysDataTransferName.FromTableID " & "JOIN ASRSysTables ToTable " & "  ON ToTable.TableID = ASRSysDataTransferName.ToTableID " & "WHERE DataTransferID = " & CStr(mlngSelectedID)
            rsTemp = DB.GetDataTable(strSQL)

            Dim objRow = rsTemp.Rows(0)


            With rsTemp
                mstrTransferName = objRow("Name").ToString
                mlngFromTableID = CInt(objRow("fromTableID"))
                mstrFromTableName = objRow("FromTableName").ToString
                mlngRecordDescExprID = CInt(objRow("RecDescExprID"))
                mlngToTableID = CInt(objRow("toTableID"))
                mstrToTableName = objRow("ToTableName").ToString
                mlngFilterID = CInt(objRow("FilterID"))
                mlngPicklistID = CInt(objRow("PicklistID"))

                mbDefinitionOwner = (LCase(Trim(_login.Username)) = LCase(Trim(objRow("UserName").ToString)))
            End With

            'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            rsTemp = Nothing

            If Not IsRecordSelectionValid() Then
                fOK = False
                Exit Sub
            End If

            strSQL = "SELECT " & "FromTable.TableName AS FromTableName, " & "FromColumn.ColumnName AS FromColumnName, " & "ToTable.TableName AS ToTableName, " & "ToColumn.ColumnName AS ToColumnName, " & "ASRSysDataTransferColumns.* " & "FROM ASRSysDataTransferColumns " & "LEFT OUTER JOIN ASRSysTables FromTable " & "  ON FromTable.TableID = ASRSysDataTransferColumns.FromTableID " & "LEFT OUTER JOIN ASRSysColumns FromColumn " & "  ON FromColumn.ColumnID = ASRSysDataTransferColumns.FromColumnID " & "LEFT OUTER JOIN ASRSysTables ToTable " & "  ON ToTable.TableID = ASRSysDataTransferColumns.ToTableID " & "LEFT OUTER JOIN ASRSysColumns ToColumn " & "  ON ToColumn.ColumnID = ASRSysDataTransferColumns.ToColumnID " & "WHERE DataTransferID = " & CStr(mlngSelectedID) & " ORDER BY ToTableID"
            rsDataTransferColumns = DB.GetDataTable(strSQL)


        Catch ex As Exception
            mstrStatusMessage = "Error reading definition"
            fOK = False

        End Try


    End Sub


    Private Sub CreateInsertTableArray()

        Dim strSQL As String
        Dim rsChildTables As DataTable

        ReDim mstrSQLTable(0)
        ReDim mlngTableViews(2, 0)

        Try

            'mstrSQLFrom remains the same all insert statements !
            mstrSQLFrom = gcoTablePrivileges.Item(mstrFromTableName).RealSource

            mstrSQLInsert = vbNullString
            mstrSQLSelect = vbNullString
            mstrSQLJoin = vbNullString
            mstrSQLWhere = vbNullString


            Call CreateChildToChildRef()
            Call BuildInsertForTable(mlngToTableID)
            If fOK = False Then
                Exit Sub
            End If


            'This will get all of the destination child table names
            'used in the definition
            strSQL = "SELECT DISTINCT " & "ASRSysDataTransferColumns.ToTableID, " & "ASRSysTables.TableName AS ToTableName " & "FROM ASRSysDataTransferColumns " & "LEFT OUTER JOIN ASRSysTables " & "  ON ASRSysTables.TableID = ASRSysDataTransferColumns.ToTableID " & "WHERE DataTransferID = " & CStr(mlngSelectedID) & " AND ToTableID <> " & CStr(mlngToTableID)
            rsChildTables = DB.GetDataTable(strSQL)

            For Each objRow As DataRow In rsChildTables.Rows
                'This is a child table so also need to insert into
                'ID column so that child references parent record
                mstrSQLInsert = "ID_" & mlngToTableID
                mstrSQLSelect = "@MaxID"
                mstrSQLJoin = vbNullString
                mstrSQLWhere = vbNullString

                Call BuildInsertForTable(CInt(objRow("toTableID")))
                If fOK = False Then
                    Exit Sub
                End If

            Next


        Catch ex As Exception
            mstrStatusMessage = "Error reading columns to transfer"
            fOK = False

        End Try

    End Sub


    Private Function GetInsertObject(lngTableID As Integer) As String

        'This sub will return the table/view which can access all
        'of the columns on the table which is in the data transfer
        'if no table/view can access all columns then vbnullstring
        'will be returned.

        Dim objTableView As TablePrivilege
        Dim sTableName As String

        fOK = True
        sTableName = "<unknown>"

        For Each objTableView In gcoTablePrivileges.Collection
            If objTableView.TableID = lngTableID Then
                sTableName = objTableView.TableName

                If objTableView.AllowInsert Then
                    If ObjectCanUpdateAllColumns(objTableView) Then
                        Return objTableView.RealSource

                    End If

                    If fOK = False Then
                        Exit Function
                    End If
                End If
            End If
        Next objTableView

        'If code reaches this point then we have looped through
        'all of the tables/views and not found any which can
        'access all of the columns on this table within the
        'definition.  Therefore we will not be able to proceed
        'with the insert command nor the data transfer
        GetInsertObject = vbNullString
        'JPD 20030411 Fault 5320
        mstrStatusMessage = "You do not have permission to insert all of the required " & "columns on the '" & sTableName & "' table."
        fOK = False

    End Function


    Private Function ObjectCanUpdateAllColumns(objTableView As TablePrivilege) As Boolean

        'This sub will check the objTableView which is passed into this sub
        'and return true/false as to whether it it can update all columns

        Dim objColumnPrivileges As CColumnPrivileges

        Try

            If objTableView.IsTable Then
                objColumnPrivileges = GetColumnPrivileges((objTableView.TableName))
            Else
                objColumnPrivileges = GetColumnPrivileges((objTableView.ViewName))
            End If

            For Each rowData As DataRow In rsDataTransferColumns.Rows

                If CInt(rowData("toTableID")) = objTableView.TableID Then

                    If objColumnPrivileges.IsValid(rowData("ToColumnName").ToString) = False Then
                        Return False

                    ElseIf objColumnPrivileges.Item(rowData("ToColumnName").ToString).AllowUpdate = False Then
                        Return False

                    End If

                End If

            Next

            Return True

        Catch ex As Exception
            mstrStatusMessage = "Error reading destination table permissions"
            fOK = False
            Return False

        End Try


    End Function


    Private Sub BuildInsertForTable(lngTableID As Integer)

        Dim sRealSource As String
        Dim strThisColumn As String
        Dim strInsertObjectName As String

        fOK = True

        'Check if we can do an insert either directly on
        'the table or through a view on the table
        strInsertObjectName = GetInsertObject(lngTableID)
        If strInsertObjectName = vbNullString Then
            Exit Sub
        End If

        Try

            ReDim mlngTableViews(2, 0)

            For Each objRow As DataRow In rsDataTransferColumns.Rows

                If CInt(objRow("toTableID")) = lngTableID Then

                    mstrSQLInsert &= IIf(mstrSQLInsert <> vbNullString, ", ", "").ToString & objRow("ToColumnName").ToString

                    If CInt(objRow("fromColumnID")) = 0 Then
                        'Either System date or free text
                        If CBool(objRow("fromSysDate")) Then
                            strThisColumn = Replace(VB6.Format(Now, "MM/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/")
                        Else
                            strThisColumn = Replace(objRow("fromText").ToString, "'", "''")
                        End If
                        strThisColumn = "'" & strThisColumn & "'"


                    ElseIf CheckDirectTableAccess(objRow("FromTableName").ToString, objRow("FromColumnName").ToString) Then
                        'Get directly from table

                        sRealSource = gcoTablePrivileges.Item(objRow("FromTableName").ToString).RealSource
                        strThisColumn = sRealSource & "." & objRow("FromColumnName").ToString

                        'Add Child table to join code
                        If sRealSource <> mstrSQLFrom Then
                            Call AddToJoinCode(0, CInt(objRow("fromTableID")), CInt(objRow("fromTableID")), sRealSource)
                        End If


                    Else
                        'Build a case statement checking access from each view
                        strThisColumn = CheckViewAccess(CInt(objRow("fromTableID")), objRow("FromTableName").ToString, objRow("FromColumnName").ToString)
                        If fOK = False Then
                            Exit Sub
                        End If

                    End If

                    mstrSQLSelect = mstrSQLSelect & IIf(mstrSQLSelect <> vbNullString, ", ", "").ToString & strThisColumn

                End If

            Next

            AddToInsertTableArray(strInsertObjectName)


        Catch ex As Exception
            mstrStatusMessage = "Error processing transfer columns"
            fOK = False

        End Try

    End Sub

    ' Returns true/false as to whether the user can directly access columm
    Private Function CheckDirectTableAccess(strTableName As String, strColumnName As String) As Boolean

        Dim objColumnPrivileges As CColumnPrivileges
        Dim bFound As Boolean

        Try

            bFound = False

            objColumnPrivileges = GetColumnPrivileges(strTableName)

            With objColumnPrivileges
                If .IsValid(strColumnName) Then
                    bFound = .Item(strColumnName).AllowSelect
                End If
            End With

            objColumnPrivileges = Nothing

        Catch ex As Exception
            mstrStatusMessage = "Error reading source table permissions"
            fOK = False

        End Try

        Return bFound

    End Function


    Private Function CheckViewAccess(lngTableID As Integer, strTableName As String, strColumnName As String) As String

        'This sub loop through all of the views on the given table and
        'builds the select case statement, adding the views to the join
        'statement where required.

        Dim objTableView As TablePrivilege
        Dim objColumnPrivileges As CColumnPrivileges
        Dim strWhereIDs As String

        CheckViewAccess = vbNullString
        strWhereIDs = vbNullString

        Try

            For Each objTableView In gcoTablePrivileges.Collection

                With objTableView

                    'Loop thru all of the objects for this table
                    'where the user has select access and the object is a view
                    If (.TableID = lngTableID) And .AllowSelect And (Not .IsTable) Then

                        ' Get the column permission for the view.
                        objColumnPrivileges = GetColumnPrivileges(.ViewName)

                        If objColumnPrivileges.IsValid(strColumnName) Then
                            If objColumnPrivileges.Item(strColumnName).AllowSelect Then

                                CheckViewAccess = CheckViewAccess & " WHEN NOT " & .ViewName & "." & strColumnName & " IS NULL THEN " & .ViewName & "." & strColumnName & vbNewLine

                                'Add View to join Code
                                ' JPD20030314 Fault 5159
                                If AddToJoinCode(1, .TableID, .ViewID, .ViewName) Then

                                    'If this view is on the base table
                                    If .TableID = mlngFromTableID Then
                                        strWhereIDs = strWhereIDs & IIf(strWhereIDs <> vbNullString, " OR ", "").ToString & .ViewName & ".ID IN (SELECT ID FROM " & mstrFromTableName & ")"
                                    End If
                                End If
                            End If
                        End If

                        'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                        objColumnPrivileges = Nothing

                    End If

                End With

            Next objTableView
            'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            objTableView = Nothing


            If CheckViewAccess <> vbNullString Then
                CheckViewAccess = "CASE" & CheckViewAccess & " ELSE NULL END"
            Else
                mstrStatusMessage = "You do not have permission to see the '" & strTableName & "." & strColumnName & "'" & vbNewLine & "column either directly or through any views." & vbNewLine
                fOK = False
                Exit Function
            End If


            ' The current user does have permission to 'read'
            ' the column through a/some view(s) on the table.
            If strWhereIDs <> vbNullString Then
                mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", "").ToString & "(" & strWhereIDs & ")"
            End If


        Catch ex As Exception
            mstrStatusMessage = "Error checking available viewing on source table"
            fOK = False

        End Try

    End Function

    Private Function AddToJoinCode(intType As Short, lngTableID As Integer, lngObjectID As Integer, strSource As String) As Boolean

        Dim iNextIndex As Integer
        Dim bAddToJoinCode As Boolean = False

        Try

            For iNextIndex = 1 To UBound(mlngTableViews, 2)
                If mlngTableViews(1, iNextIndex) = intType And mlngTableViews(2, iNextIndex) = lngObjectID Then
                    Return False
                End If
            Next iNextIndex

            bAddToJoinCode = True

            ' if this column is not from the base table then it must be from a parent
            ' table, therefore include it in the join code
            iNextIndex = UBound(mlngTableViews, 2) + 1
            ReDim Preserve mlngTableViews(2, iNextIndex)
            mlngTableViews(1, iNextIndex) = intType
            mlngTableViews(2, iNextIndex) = lngObjectID

            ' The table has not yet been added to the join code, and it is
            ' not the base table so add it to the array and the join code.
            If lngTableID <> mlngFromTableID Then
                'If its a child table then make sure that
                'we reference the parent ID column and do an inner join
                mstrSQLJoin = mstrSQLJoin & " INNER JOIN " & strSource & " ON " & mstrFromTableName & ".ID" & " = " & strSource & ".ID_" & CStr(mlngFromTableID)
            Else
                'Join view of parent and left outer join
                mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & strSource & " ON " & mstrFromTableName & ".ID" & " = " & strSource & ".ID"

            End If

        Catch ex As Exception
            mstrStatusMessage = "Error joining source tables"
            fOK = False

        End Try

        Return bAddToJoinCode

    End Function


    Private Sub AddToInsertTableArray(strInsertObjectName As String)

        Dim intIndex As Integer


        Try

            If mstrSQLTable(0) = vbNullString Then
                intIndex = 0
            Else
                intIndex = UBound(mstrSQLTable) + 1
                ReDim Preserve mstrSQLTable(intIndex)
            End If


            mstrSQLTable(intIndex) = "INSERT " & strInsertObjectName & vbNewLine & "(" & mstrSQLInsert & ")" & vbNewLine

            mstrSQLTable(intIndex) = mstrSQLTable(intIndex) & "SELECT " & mstrSQLSelect & vbNewLine & "FROM " & mstrSQLFrom & vbNewLine

            If mstrSQLJoin <> vbNullString Then
                mstrSQLTable(intIndex) = mstrSQLTable(intIndex) & mstrSQLJoin & vbNewLine
            End If

            mstrSQLTable(intIndex) = mstrSQLTable(intIndex) & " WHERE " & mstrSQLWhere & IIf(mstrSQLWhere <> vbNullString, " AND ", "").ToString & mstrSQLFrom & ".ID = @RecordID" & vbNewLine & vbNewLine


            If intIndex = 0 Then
                'If its the first table then select the max id from the table afterwards
                mstrGetMaxID = strInsertObjectName
            End If

        Catch ex As Exception
            mstrStatusMessage = "Error building SQL insert statement"
            fOK = False

        End Try

    End Sub


    Private Sub CreateStoredProcedure()

        Dim strSQL As String
        Dim intCount As Integer

        strSQL = "/* ------------------------------------- */" & vbNewLine & "/* OpenHR Data Transfer stored procedure. */" & vbNewLine & "/* ------------------------------------- */" & vbNewLine & "CREATE PROCEDURE " & mstrProcedureName & "(@RecordID int)" & vbNewLine & "AS" & vbNewLine & "BEGIN" & vbNewLine & vbNewLine


        Try

            strSQL = strSQL & "IF NOT EXISTS(SELECT ID FROM " & mstrSQLFrom & " WHERE ID = @RecordID)" & vbNewLine & "BEGIN" & vbNewLine & "  RAISERROR('Record not found', 16, 1)" & vbNewLine & "  RETURN 1" & vbNewLine & "END" & vbNewLine & vbNewLine

            If UBound(mstrSQLTable) > 0 Then
                'More then one table to insert to

                strSQL = strSQL & "DECLARE @MaxID int" & vbNewLine & vbNewLine & "BEGIN TRANSACTION" & vbNewLine & vbNewLine

                ' JPD20030314 Fault 5159
                strSQL = strSQL & mstrSQLTable(0) & vbNewLine & vbNewLine & "/* -------------------------------------------------- */" & vbNewLine & "/* Get max id from primary table to ensure that child */" & vbNewLine & "/* tables correctly reference the parent table        */" & vbNewLine & "/* -------------------------------------------------- */" & vbNewLine & "EXEC spASRMaxID " & Trim(Str(mlngToTableID)) & ", @MaxID OUTPUT" & vbNewLine
                '"SELECT @MaxID = MAX(ID) FROM " & mstrGetMaxID & vbNewLine

                For intCount = 1 To UBound(mstrSQLTable)
                    strSQL = strSQL & mstrSQLTable(intCount) & vbNewLine & vbNewLine
                Next

                strSQL = strSQL & "COMMIT TRANSACTION" & vbNewLine & vbNewLine

            Else
                'Simple insert on a single table
                strSQL = strSQL & mstrSQLTable(0) & vbNewLine & vbNewLine

            End If

            strSQL = strSQL & "END"
            DB.ExecuteSql(strSQL)

        Catch ex As Exception
            mstrStatusMessage = "Error creating stored procedure"
            fOK = False

        End Try

    End Sub


    Private Sub ProcessRecords()

        Dim rsRecords As DataTable
        Dim strSQL As String
        Dim strRecordError As String

        rsRecords = GetRecordIDs()

        If fOK = False Then
            Exit Sub
        End If

        'Run the stored procedure for each record id
        mlngSuccessCount = 0
        mlngFailCount = 0
        For Each objRow As DataRow In rsRecords.Rows

            Try

                Dim prmNewID = New SqlParameter("@piNewRecordID", SqlDbType.Int) With {.Value = 0, .Direction = ParameterDirection.Output}
                strSQL = "EXEC " & mstrProcedureName & " " & objRow("ID").ToString

                DB.ExecuteSP("sp_ASRInsertNewRecord" _
                    , prmNewID _
                    , New SqlParameter("@psInsertString", SqlDbType.NVarChar, -1) With {.Value = strSQL})

                mlngSuccessCount = mlngSuccessCount + 1

                If mbLoggingDTSuccess Then
                    Logs.AddDetailEntry(GetRecordDesc(CInt(objRow("ID"))) & " transferred successfully")
                End If


            Catch ex As Exception

                strRecordError = GetRecordDesc(CInt(objRow("ID"))) & vbNewLine & vbNewLine & ex.Message
                Call Logs.AddDetailEntry(strRecordError)
                mlngFailCount += 1

            End Try

        Next

    End Sub


    Private Function GetRecordIDs() As DataTable

        Dim objTableView As TablePrivilege
        Dim rsTemp As DataTable
        Dim strSQL As String
        Dim lngLoop As Integer
        Dim strPicklistFilterIDs As String
        Dim blnAllowed As Boolean

        Try


            'Set up an array of record ids to be updated
            Dim mlngRecordIDs(0) As Integer

            strPicklistFilterIDs = GetPicklistFilterSelect()

            If fOK = False Then
                Exit Function
            End If


            'We need to build a string which will select the ID from all of the views
            'that we have select permission on
            strSQL = vbNullString
            For Each objTableView In gcoTablePrivileges.Collection

                blnAllowed = (objTableView.TableID = mlngFromTableID And objTableView.AllowSelect)

                If blnAllowed Then
                    strSQL = strSQL & IIf(strSQL <> vbNullString, vbNewLine & "UNION" & vbNewLine, "").ToString & "SELECT ID FROM " & objTableView.RealSource & IIf(strPicklistFilterIDs <> vbNullString, " WHERE ID IN (" & strPicklistFilterIDs & ")", "").ToString
                End If

            Next objTableView

            'Now get distinct IDs from the base table (as there couuld be duplicates in the above sql!)
            strSQL = "SELECT ID FROM " & mstrSQLFrom & IIf(strSQL <> vbNullString, " WHERE ID IN (" & strSQL & ")", "").ToString

            ' Create any UDFs
            fOK = General.UDFFunctions(mastrUDFsRequired, True)

            Return DB.GetDataTable(strSQL)

        Catch ex As Exception
            mstrStatusMessage = "Error selecting records"
            fOK = False

        End Try

    End Function

    Private Sub TidyUpAndExit()

        General.DropUniqueSQLObject(mstrProcedureName, 4)

    End Sub


    Private Sub CreateChildToChildRef()

        Dim rsTemp As DataTable
        Dim strSQL As String

        strSQL = "SELECT ParentID FROM ASRSysRelations " & "WHERE ParentID IN " & "(SELECT ParentID FROM ASRSysRelations " & " WHERE ChildID = " & CStr(mlngFromTableID) & ") AND ChildID = " & CStr(mlngToTableID)

        rsTemp = DB.GetDataTable(strSQL)

        mstrSQLInsert = vbNullString

        For Each objRow As DataRow In rsTemp.Rows
            mstrSQLInsert = mstrSQLInsert & IIf(mstrSQLInsert <> vbNullString, ", ", "").ToString & "ID_" & CStr(CInt(objRow("ParentID").ToString))
        Next

        mstrSQLSelect = mstrSQLInsert

    End Sub

    Private Sub CheckIfTransferIsValid()

        Dim rsTemp As DataTable
        Dim strSQL As String
        Dim strSQL1 As String
        Dim strSQL2 As String
        Dim strSQL3 As String
        Dim strTablesInDef As String
        Dim strColumnsInDef As String
        Dim strErrorColumns As String = ""

        Const SQLWhereMandatoryColumn =
            "(Rtrim(DefaultValue) = '' OR (Rtrim(DefaultValue) = '__/__/____') and DataType = 11)" &
            " AND Convert(int,isnull(dfltValueExprID,0)) = 0" &
            " AND CalcExprID = 0" &
            " AND Mandatory = '1'" &
            " AND ColumnType <> 4 "


        strTablesInDef = "SELECT DISTINCT ASRSysDataTransferColumns.ToTableID " & "FROM ASRSysDataTransferColumns " & "WHERE ASRSysDataTransferColumns.DataTransferID = " & CStr(mlngSelectedID)

        strColumnsInDef = "SELECT ASRSysDataTransferColumns.ToColumnID " & "FROM ASRSysDataTransferColumns " & "WHERE ASRSysDataTransferColumns.DataTransferID = " & CStr(mlngSelectedID)


        'This will retreive all of the read only
        'destination columns in the data transfer
        strSQL1 = "SELECT ASRSysTables.TableName+'.'+ASRSysColumns.ColumnName as 'TableColumn', " & "'Read Only' as Reason " & "FROM ASRSysDataTransferColumns " & "JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysDataTransferColumns.ToColumnID " & "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & "WHERE ASRSysDataTransferColumns.DataTransferID = " & CStr(mlngSelectedID) & " AND ASRSysColumns.ReadOnly <> 0"


        'This will get all of the mandatory columns
        'which have not been included in the definition

        'MH20000814
        'Allow save if mandatory ommitted if it has a default value
        'This is to get around the staff number on a applicants to personnel transfer

        'MH20000904
        'Allow save if mandatory ommitted and it is a calculated column

        '******************************************************************************
        ' TM20010719 Fault 2242 - ColumnType <> 4 clause added to ignore all linked   *
        ' columns. (It doesn't need to validate the linked columns because this is    *
        ' done using the Vaidate SP.                                                  *
        '******************************************************************************

        strSQL2 = "SELECT ASRSysTables.TableName+'.'+ASRSysColumns.ColumnName as 'TableColumn', " & "'Mandatory' as Reason " & "FROM ASRSysColumns " &
            "JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID " & "WHERE ASRSysColumns.TableID IN (" & strTablesInDef & ") " &
            " AND ASRSysColumns.ColumnID NOT IN (" & strColumnsInDef & ") " & " AND " & SQLWhereMandatoryColumn
        '" AND Mandatory = '1'" & _
        '" AND CalcExprID = 0 " & _
        '" AND ColumnType <> 4 " & _
        '" AND Rtrim(DefaultValue) = '' AND Convert(int,dfltValueExprID) = 0 "


        'This will ensure matching data types and
        'compatable column sizes
        strSQL3 = "SELECT FromTable.TableName+'.'+FromColumn.ColumnName+' / '+ToTable.TableName+'.'+ToColumn.ColumnName as 'TableColumn', " &
            "'Type or Size' as Reason " & "FROM ASRSysDataTransferColumns " & "JOIN ASRSysColumns FromColumn ON FromColumn.ColumnID = ASRSysDataTransferColumns.FromColumnID " &
            "JOIN ASRSysTables FromTable ON FromTable.TableID = FromColumn.TableID " & "JOIN ASRSysColumns ToColumn ON ToColumn.ColumnID = ASRSysDataTransferColumns.ToColumnID " &
            "JOIN ASRSysTables ToTable ON ToTable.TableID = ToColumn.TableID " & "WHERE ASRSysDataTransferColumns.DataTransferID = " & CStr(mlngSelectedID) &
            "  AND (FromColumn.DataType <> ToColumn.DataType OR " & "FromColumn.Size > ToColumn.Size OR " & "FromColumn.Decimals > ToColumn.Decimals)"


        strSQL = strSQL1 & " UNION " & strSQL2 & " UNION " & strSQL3 & " ORDER BY 'TableColumn'"

        rsTemp = DB.GetDataTable(strSQL)

        For Each objRow As DataRow In rsTemp.Rows
            strErrorColumns = strErrorColumns & objRow("Reason").ToString & " :" & vbTab & objRow("TableColumn").ToString & vbNewLine
        Next

        If strErrorColumns.Length > 0 Then
            mstrStatusMessage = "Unable to run this Data Transfer as the following " &
                        "column definitions have changed:" & vbNewLine & vbNewLine &
                        strErrorColumns
        End If

    End Sub


    Private Function GetPicklistFilterSelect() As String

        Dim rsTemp As DataTable

        GetPicklistFilterSelect = vbNullString

        'If mlngSingleRecordID > 0 Then
        '  GetPicklistFilterSelect = CStr(mlngSingleRecordID)

        If mstrPicklistFilterIDs <> "" Then
            GetPicklistFilterSelect = mstrPicklistFilterIDs


        ElseIf mlngPicklistID > 0 Then

            mstrStatusMessage = IsPicklistValid(mlngPicklistID)

            'Get List of IDs from Picklist
            rsTemp = DB.GetDataTable("sp_ASRGetPickListRecords " & mlngPicklistID)

            fOK = rsTemp.Rows.Count > 0

            If Not fOK Then
                mstrStatusMessage = "The base table picklist contains no records."
            Else
                For Each objRow As DataRow In rsTemp.Rows
                    GetPicklistFilterSelect = GetPicklistFilterSelect & IIf(Len(GetPicklistFilterSelect) > 0, ", ", "").ToString & objRow(0).ToString
                Next

            End If

        ElseIf mlngFilterID > 0 Then

            mstrStatusMessage = IsFilterValid(mlngFilterID)

            'Get list of IDs from Filter
            fOK = FilteredIDs(mlngFilterID, GetPicklistFilterSelect, mastrUDFsRequired)

            ' Generate any UDFs that are used in this filter
            If fOK Then
                General.UDFFunctions(mastrUDFsRequired, True)
            End If

            If Not fOK Then
                ' Permission denied on something in the filter.
                mstrStatusMessage = "You do not have permission to use the '" & General.GetFilterName(mlngFilterID) & "' filter."
            End If

        End If

    End Function


    Private Function ValidSQL(strSQL As String, ByRef strErrorMsg As String) As Boolean

        Try
            DB.ExecuteSql(strSQL)

        Catch ex As Exception
            strErrorMsg = ex.Message
            Return False

        End Try

        Return True

    End Function


    Private Function GetRecordDesc(lngRecordID As Integer) As String

        Dim prmRecordDesc = New SqlParameter("psRecDesc", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

        Dim unKnownVar1 As Integer = 0
        Dim unKnownVar2 As Integer = 0
        Dim unKnownVar3 As Integer = 0

        DB.ExecuteSP("sp_ASRIntGetRecordDescription" _
                            , New SqlParameter("piTableID", SqlDbType.Int) With {.Value = unKnownVar1} _
                            , New SqlParameter("piRecordID", SqlDbType.Int) With {.Value = lngRecordID} _
                            , New SqlParameter("piParentTableID", SqlDbType.Int) With {.Value = unKnownVar2} _
                            , New SqlParameter("piParentRecordID", SqlDbType.Int) With {.Value = unKnownVar3} _
                            , prmRecordDesc)

        Return prmRecordDesc.Value.ToString

    End Function

End Class