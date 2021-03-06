﻿Option Strict Off
Option Explicit On
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata

Friend Class clsExportRUN

    ' To hold Properties
    Private mlngExportID As Integer
    Private mstrErrorString As String
    Private mblnEnableSQLTable As Boolean
    Private mbNoRecords As Boolean

    ' String to hold the temp table name
    Private mstrTempTableName As String
    Private mstrArrayHeader(,) As String
    Private mstrArrayData(,) As String
    Private mstrArrayFooter(,) As String

    ' Variables to store definition
    Private mstrExportName As String
    Private mstrExportDescription As String
    Private mlngExportBaseTable As Integer
    Private mstrExportBaseTableName As String
    Private mlngExportAllRecords As Integer
    Private mlngExportPickListID As Integer
    Private mlngExportFilterID As Integer
    Private mlngExportParent1Table As Integer
    Private mstrExportParent1TableName As String
    Private mlngExportParent1FilterID As Integer
    Private mlngExportParent2Table As Integer
    Private mstrExportParent2TableName As String
    Private mlngExportParent2FilterID As Integer
    Private mlngExportChildTable As Integer
    Private mstrExportChildTableName As String
    Private mlngExportChildFilterID As Integer
    Private mlngExportChildMaxRecords As Integer
    'Private mstrExportOutputType As String
    'Private mstrOutputFilename As String
    Private mblnExportQuotes As Boolean
    Private mblnStripDelimiter As Boolean
    'Private mblnExportHeader As Boolean
    Private mintExportHeader As Short
    Private mstrExportHeaderText As String
    Private mintExportFooter As Short
    Private mstrExportFooterText As String
    Private mstrExportDelimiter As String
    Private mstrExportOtherDelimiter As String
    Private mstrExportActualDelimiter As String
    Private mstrExportDateFormat As String
    Private mstrExportDateSeparator As String
    'Private mstrExportDateYearDigits As String
    Private mlngExportRecordCount As Integer

    'Private mbAppendToFile As Boolean                   ' Append export to existing file
    Private mblnHeader As Boolean
    Private mblnFooter As Boolean
    Private mbForceHeader As Boolean ' Always force header if no records found
    Private mbOmitHeader As Boolean ' Omit header if appending to file
    Private mbAuditChangesOnly As Boolean

    Private mbSplitFile As Boolean
    Private mlngSplitFileSize As Integer

    Private mlngOutputFormat As Integer
    Private mblnOutputSave As Boolean
    Private mlngOutputSaveExisting As Integer
    'Private mlngOutputSaveFormat As Long
    Private mblnOutputEmail As Boolean
    Private mlngOutputEmailAddr As Integer
    Private mstrOutputEmailSubject As String
    Private mstrOutputEmailAttachAs As String
    'Private mlngOutputEmailFileFormat As Long
    Private mstrOutputFileName As String

    Private mlngExportParent1AllRecords As Integer
    Private mlngExportParent1PickListID As Integer
    Private mlngExportParent2AllRecords As Integer
    Private mlngExportParent2PickListID As Integer

    ' Recordsets to store the definition/column information and the final output data
    Private mrstExportDetails As New ADODB.Recordset
    Private mrstExportOutput As New ADODB.Recordset

    ' Strings to hold the SQL statement
    Private mstrSQLSelect As String
    Private mstrSQLFrom As String
    Private mstrSQLJoin As String
    Private mstrSQLWhere As String
    Private mstrSQLOrderBy As String
    Private mstrSQL As String
    Private mstrOnlyChangesFilter As String

    Private mstrTransformFile As String
    Private mstrXMLDataNodeName As String
    Private mstrXSDFilename As String
    Private mbPreserveTransformPath As Boolean
    Private mbPreserveXSDPath As Boolean
    Private mbSplitXMLIntoFiles As Boolean

    ' Data access classes
    Private mclsData As clsDataAccess
    Private mclsGeneral As clsGeneral

    ' Array holding the columns to sort the report by
    Private mvarSortOrder(,) As Object

    ' Array to hold the columns used in the export
    Dim mvarColDetails(,) As Object

    ' TableViewsGuff
    Private mstrRealSource As String
    Private mstrBaseTableRealSource As String
    Private mlngTableViews(,) As Integer
    Private mstrViews() As String
    Private mobjTableView As TablePrivilege
    Private mobjColumnPrivileges As CColumnPrivileges

    ' Batch mode ?
    'Private gblnBatchMode As Boolean

    Private mblnUserCancelled As Boolean
    Private mblnNoRecords As Boolean

    ' CMG File Code
    Private mstrExportFileCode As String
    Private mlngCMGRecordIdentifier As Integer
    Private mbUpdateAuditLog As Boolean
    Private mdLastSuccessfulOutput As Date

    Private mdExportCreateDate As Date

    Private Enum CMGFields
        NewValue = 0 ' "NewValue"
        DateTime = 1 ' "DateTimeStamp"
        ColumnID = 2 ' "ColumnID"
    End Enum

    Private mbDefinitionOwner As Boolean
    Private mbLoggingExportSuccess As Boolean

    Private mlSuccessfulRecords As Integer

    ' Array holding the User Defined functions that are needed for this report
    Private mastrUDFsRequired() As String


    Public ReadOnly Property UserCancelled() As Boolean
        Get
            UserCancelled = mblnUserCancelled
        End Get
    End Property


    Public Property ExportID() As Integer
        Get
            ExportID = mlngExportID
        End Get
        Set(ByVal Value As Integer)
            mlngExportID = Value
        End Set
    End Property


    Public Property ErrorString() As String
        Get
            ErrorString = mstrErrorString
        End Get
        Set(ByVal Value As String)
            mstrErrorString = Value
        End Set
    End Property

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()

        ' Purpose : Sets references to other classes and redimensions arrays
        '           used for table usage information
        ' Input   : None
        ' Output  : None

        mclsData = New DataMgr.clsDataAccess
        mclsGeneral = New DataMgr.clsGeneral
        ReDim mvarSortOrder(2, 0)
        'ReDim mvarColDetails(6, 0)
        '  ReDim mvarColDetails(10, 0)
        '  ReDim mvarColDetails(15, 0)
        'NPG20071217 Fault 12867
        ' ReDim mvarColDetails(16, 0)
        'NPG20080617 Suggestion S000816
        ReDim mvarColDetails(17, 0)
        ReDim mlngTableViews(2, 0)
        ReDim mstrViews(0)

        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mblnEnableSQLTable = (GetSystemSetting("Output", "ExportToSQLTable", 0) = 1)

    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Terminate_Renamed()

        ' Purpose : Clears references to other classes.
        ' Input   : None
        ' Output  : None

        'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mclsData = Nothing
        'UPGRADE_NOTE: Object mclsGeneral may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mclsGeneral = Nothing

    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub

    Private Function IsRecordSelectionValid() As Boolean

        Dim sSQL As String
        Dim lCount As Integer
        Dim rsTemp As ADODB.Recordset
        Dim iResult As RecordSelectionValidityCodes

        ' Filter
        If mlngExportFilterID > 0 Then
            iResult = ValidateRecordSelection(RecordSelectionType.Filter, mlngExportFilterID)
            Select Case iResult
                Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                    mstrErrorString = "The base table filter used in this definition has been deleted by another user."
                Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                    mstrErrorString = "The base table filter used in this definition is invalid."
                Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                    If Not gfCurrentUserIsSysSecMgr Then
                        mstrErrorString = "The base table filter used in this definition has been made hidden by another user."
                    End If
            End Select
        ElseIf mlngExportPickListID > 0 Then
            iResult = ValidateRecordSelection(RecordSelectionType.Picklist, mlngExportPickListID)
            Select Case iResult
                Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                    mstrErrorString = "The base table picklist used in this definition has been deleted by another user."
                Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                    mstrErrorString = "The base table picklist used in this definition is invalid."
                Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                    If Not gfCurrentUserIsSysSecMgr Then
                        mstrErrorString = "The base table picklist used in this definition has been made hidden by another user."
                    End If
            End Select
        End If

        If Len(mstrErrorString) = 0 Then
            If mlngExportParent1FilterID > 0 Then
                iResult = ValidateRecordSelection(RecordSelectionType.Filter, mlngExportParent1FilterID)
                Select Case iResult
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                        mstrErrorString = "The first parent table filter used in this definition has been deleted by another user."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                        mstrErrorString = "The first parent table filter used in this definition is invalid."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                        If Not gfCurrentUserIsSysSecMgr Then
                            mstrErrorString = "The first parent table filter used in this definition has been made hidden by another user."
                        End If
                End Select
            ElseIf mlngExportParent1PickListID > 0 Then
                iResult = ValidateRecordSelection(RecordSelectionType.Picklist, mlngExportParent1PickListID)
                Select Case iResult
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                        mstrErrorString = "The first parent table picklist used in this definition has been deleted by another user."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                        mstrErrorString = "The first parent table picklist used in this definition is invalid."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                        If Not gfCurrentUserIsSysSecMgr Then
                            mstrErrorString = "The first parent table picklist used in this definition has been made hidden by another user."
                        End If
                End Select
            End If
        End If

        If Len(mstrErrorString) = 0 Then
            If mlngExportParent2FilterID > 0 Then
                iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_FILTER, mlngExportParent2FilterID)
                Select Case iResult
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                        mstrErrorString = "The second parent table filter used in this definition has been deleted by another user."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                        mstrErrorString = "The second parent table filter used in this definition is invalid."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                        If Not gfCurrentUserIsSysSecMgr Then
                            mstrErrorString = "The second parent table filter used in this definition has been made hidden by another user."
                        End If
                End Select
            ElseIf mlngExportParent2PickListID > 0 Then
                iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_PICKLIST, mlngExportParent2PickListID)
                Select Case iResult
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                        mstrErrorString = "The second parent table picklist used in this definition has been deleted by another user."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                        mstrErrorString = "The second parent table picklist used in this definition is invalid."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                        If Not gfCurrentUserIsSysSecMgr Then
                            mstrErrorString = "The second parent table picklist used in this definition has been made hidden by another user."
                        End If
                End Select
            End If
        End If

        If Len(mstrErrorString) = 0 Then
            If mlngExportChildFilterID > 0 Then
                iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_FILTER, mlngExportChildFilterID)
                Select Case iResult
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                        mstrErrorString = "The child table filter used in this definition has been deleted by another user."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                        mstrErrorString = "The child table filter used in this definition is invalid."
                    Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                        If Not gfCurrentUserIsSysSecMgr Then
                            mstrErrorString = "The child table filter used in this definition has been made hidden by another user."
                        End If
                End Select
            End If
        End If


        '******* Check calculations for hidden/deleted elements *******

        If Len(mstrErrorString) = 0 Then
            sSQL = "SELECT * FROM ASRSYSExportDetails " & "WHERE ExportID = " & mlngExportID & " AND LOWER(Type) = 'x' "

            rsTemp = datGeneral.GetRecords(sSQL)
            With rsTemp
                If Not (.EOF And .BOF) Then
                    .MoveFirst()
                    Do Until .EOF
                        iResult = ValidateCalculation(.Fields("ColExprID").Value)
                        Select Case iResult
                            Case RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
                                mstrErrorString = "A calculation used in this definition has been deleted by another user."
                            Case RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
                                mstrErrorString = "A calculation used in this definition is invalid."
                            Case RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
                                If Not gfCurrentUserIsSysSecMgr Then
                                    mstrErrorString = "A calculation used in this definition has been made hidden by another user."
                                End If
                        End Select

                        If Len(mstrErrorString) > 0 Then
                            Exit Do
                        End If

                        .MoveNext()
                    Loop
                End If
            End With

            'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            rsTemp = Nothing
        End If

        IsRecordSelectionValid = (Len(mstrErrorString) = 0)

    End Function


    Public Function RunExport() As Boolean

        ' Purpose : This function is called from frmDefsel and Batch Jobs.
        '           It is the main function which runs the export.
        ' Input   : bSilent, Boolean, Suppress COAMsgBoxs ? (ie, for batch jobs)
        ' Output  : True/False Success

        On Error GoTo RunExport_ERROR

        Dim fOK As Boolean
        Dim sToday As String

        mdExportCreateDate = Now
        mstrErrorString = vbNullString
        mblnNoRecords = False

        fOK = True
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        ' JDM - 07/10/01 - Fault 2644 - Cannot run CMG export if security is removed
        If fOK Then fOK = GetExportDefinition()

        'UPGRADE_WARNING: Couldn't resolve default property of object GetUserSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mbLoggingExportSuccess = CBool(GetUserSetting("LogEvents", "Export_Success", False))

        If fOK Then
            With gobjProgress
                '.AviFile = App.Path & "\videos\export.avi"
                .AVI = COAProgress.AVIType.dbText
                .MainCaption = "Export"
                If gblnBatchMode = False Then
                    .NumberOfBars = 1
                    .Caption = "Export"
                    .Time = False
                    .Cancel = True
                    '.Bar1MaxValue = 9
                    .OpenProgress()
                Else
                    .ResetBar2()
                    '.Bar2MaxValue = 9
                End If
            End With


            'If mstrExportOutputType = "C" And Not datGeneral.SystemPermission("CMG", "CMGRUN") Then
            If mlngOutputFormat = modEnums.OutputFormats.fmtCMGFile And Not datGeneral.SystemPermission("CMG", "CMGRUN") Then
                mstrErrorString = vbCrLf & "You do not have permission to run a CMG export" & vbCrLf & "Please contact your system administrator"
                fOK = False
            End If

            If gblnBatchMode Then
                gobjProgress.Bar2Caption = "Export : " & mstrExportName
            Else
                gobjProgress.Bar1Caption = "Export : " & mstrExportName
            End If

            gobjEventLog.AddHeader(modEnums.EventLog_Type.eltExport, mstrExportName)
        End If

        If fOK Then fOK = GetDetailsRecordsets()
        If fOK Then fOK = GenerateSQL()
        If fOK Then fOK = AddTempTableToSQL()
        If fOK Then fOK = MergeSQLStrings()
        If fOK Then fOK = UDFFunctions(mastrUDFsRequired, True)
        If fOK Then fOK = ExecuteSql()
        If fOK Then fOK = UDFFunctions(mastrUDFsRequired, False)
        If fOK Then fOK = CheckRecordSet()
        If fOK Then fOK = ExportData()

        'If gobjProgress.Visible Then gobjProgress.UpdateProgress gblnBatchMode

        If Not gblnBatchMode And gobjProgress.Visible Then gobjProgress.CloseProgress()

        Call UtilUpdateLastRun(modEnums.UtilityType.utlExport, mlngExportID)


        'MH20000705 Fault 519
        'No records is reported as a failure !

        'mlngExportRecordCount

        If mblnNoRecords Then
            gobjEventLog.ChangeHeaderStatus(modEnums.EventLog_Status.elsSuccessful, mlngExportRecordCount, 0)
            gobjEventLog.AddDetailEntry(mstrErrorString)
            mstrErrorString = "Completed successfully." & vbCrLf & mlngExportRecordCount & " record(s) exported." & vbCrLf & mstrErrorString
            fOK = True

        ElseIf fOK Then
            gobjEventLog.ChangeHeaderStatus(modEnums.EventLog_Status.elsSuccessful, mlngExportRecordCount, 0)
            mstrErrorString = "Completed successfully." & vbCrLf & mlngExportRecordCount & " record(s) exported."

            sToday = "convert(datetime, '" & Replace(VB6.Format(Now, "MM/dd/yyyy hh:mm:ss"), UI.GetSystemDateSeparator, "/") & "')"
            mstrSQL = "UPDATE ASRSysExportName SET LastSuccessfulOutput = " & sToday & " WHERE ID = " & mlngExportID
            mclsData.ExecuteSql(mstrSQL)

        ElseIf mblnUserCancelled Then
            gobjEventLog.ChangeHeaderStatus(modEnums.EventLog_Status.elsCancelled, mlSuccessfulRecords, 0)
            mstrErrorString = "Cancelled by user." & vbCrLf & mlSuccessfulRecords & " record(s) exported."
        Else
            'Only details records for failures !
            gobjEventLog.AddDetailEntry(mstrErrorString)
            gobjEventLog.ChangeHeaderStatus(modEnums.EventLog_Status.elsFailed, mlSuccessfulRecords, 0)
            mstrErrorString = "Failed." & vbCrLf & vbCrLf & mstrErrorString '& vbCrLf & vbCrLf & mlSuccessfulRecords & " record(s) exported."
        End If

        mstrErrorString = "Export : '" & mstrExportName & "' " & mstrErrorString

        If Not gblnBatchMode Then
            COAMsgBox(mstrErrorString, IIf(fOK, MsgBoxStyle.Information, MsgBoxStyle.Exclamation) + MsgBoxStyle.OkOnly, "Export")
        End If

        If fOK = True Then fOK = ClearUp() Else ClearUp()

        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        RunExport = fOK

        Exit Function

RunExport_ERROR:

        fOK = False
        mstrErrorString = "Error Whilst Running Export." & vbCrLf & "(" & Err.Description & ")"
        Resume Next

    End Function

    Private Function AddTempTableToSQL() As Boolean

        ' Purpose : This function retrieves a unique temp table name and
        '           inserts it into the SQL Select statement
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo AddTempTableToSQL_ERROR

        'gobjProgress.UpdateProgress gblnBatchMode

        ' If user cancels the export, abort
        If gobjProgress.Cancelled Then
            mblnUserCancelled = True
            AddTempTableToSQL = False
            Exit Function
        End If

        mstrTempTableName = datGeneral.UniqueSQLObjectName("ASRSysTempExport", 3)

        mstrSQLSelect = mstrSQLSelect & " INTO " & "[" & mstrTempTableName & "]"

        AddTempTableToSQL = True
        Exit Function

AddTempTableToSQL_ERROR:

        mstrErrorString = "Error whilst retrieving unique temp table name." & vbCrLf & "(" & Err.Description & ")"
        AddTempTableToSQL = False

    End Function

    Private Function MergeSQLStrings() As Boolean

        ' Purpose : This function merges all the SQL string variables
        '           into one long string
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo MergeSQLStrings_ERROR

        'gobjProgress.UpdateProgress gblnBatchMode

        ' If user cancels the export, abort
        If gobjProgress.Cancelled Then
            mblnUserCancelled = True
            MergeSQLStrings = False
            Exit Function
        End If

        mstrSQL = mstrSQLSelect & " FROM " & mstrSQLFrom & IIf(Len(mstrSQLJoin) = 0, "", " " & mstrSQLJoin) & IIf(Len(mstrSQLWhere) = 0, "", " " & mstrSQLWhere) & " " & mstrSQLOrderBy

        MergeSQLStrings = True
        Exit Function

MergeSQLStrings_ERROR:

        mstrErrorString = "Error whilst merging SQL string components." & vbCrLf & "(" & Err.Description & ")"
        MergeSQLStrings = False

    End Function

    Private Function ExecuteSql() As Boolean

        ' Purpose : This function executes the SQL string
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo ExecuteSQL_ERROR

        'gobjProgress.UpdateProgress gblnBatchMode

        ' If user cancels the export, abort
        If gobjProgress.Cancelled Then
            mblnUserCancelled = True
            ExecuteSql = False
            Exit Function
        End If

        'COAMsgBox "Recordset will be generated from : " & vbCrLf & mstrSQL
        mclsData.ExecuteSql(mstrSQL)

        ExecuteSql = True
        Exit Function

ExecuteSQL_ERROR:

        mstrErrorString = "Error whilst executing SQL statement." & vbCrLf & "(" & Err.Description & ")"
        ExecuteSql = False

    End Function

    Private Function GetExportDefinition() As Boolean

        ' Purpose : This function retrieves the basic definition details
        '           and stores it in module level variables
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo GetExportDefinition_ERROR

        Dim prstTemp_Definition As ADODB.Recordset
        Dim pstrSQL As String
        Dim rsTemp As ADODB.Recordset

        '  gobjProgress.UpdateProgress gblnBatchMode
        '
        '  ' If user cancels the export, abort
        '  If gobjProgress.Cancelled Then
        '    mblnUserCancelled = True
        '    GetExportDefinition = False
        '    Exit Function
        '  End If

        pstrSQL = "SELECT * FROM AsrSysExportName " & "WHERE ID = " & mlngExportID & " "

        prstTemp_Definition = mclsData.OpenRecordset(pstrSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        With prstTemp_Definition

            If .BOF And .EOF Then
                GetExportDefinition = False
                mstrErrorString = "Could not find specified Export definition !"
                prstTemp_Definition.Close()
                'UPGRADE_NOTE: Object prstTemp_Definition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                prstTemp_Definition = Nothing
                Exit Function
            End If


            mstrExportName = .Fields("Name").Value
            mstrExportDescription = .Fields("Description").Value
            mlngExportBaseTable = .Fields("BaseTable").Value
            mstrExportBaseTableName = datGeneral.GetTableName(mlngExportBaseTable)
            mlngExportAllRecords = .Fields("AllRecords").Value
            mlngExportPickListID = .Fields("picklist").Value
            mlngExportFilterID = .Fields("Filter").Value
            mlngExportParent1Table = .Fields("parent1table").Value
            mstrExportParent1TableName = datGeneral.GetTableName(mlngExportParent1Table)
            mlngExportParent1FilterID = .Fields("parent1filter").Value
            mlngExportParent2Table = .Fields("parent2table").Value
            mstrExportParent2TableName = datGeneral.GetTableName(mlngExportParent2Table)
            mlngExportParent2FilterID = .Fields("parent2filter").Value
            mlngExportChildTable = .Fields("ChildTable").Value
            mstrExportChildTableName = datGeneral.GetTableName(mlngExportChildTable)
            mlngExportChildFilterID = .Fields("childFilter").Value
            mlngExportChildMaxRecords = .Fields("ChildMaxRecords").Value
            'mstrExportOutputType = !outputtype
            'mstrExportOutputName = !outputname
            mblnExportQuotes = .Fields("Quotes").Value
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mblnStripDelimiter = IIf(IsDBNull(.Fields("StripDelimiterFromData").Value), False, .Fields("StripDelimiterFromData").Value)

            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mlngSplitFileSize = IIf(IsDBNull(.Fields("SplitFileSize").Value), 0, .Fields("SplitFileSize").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mbSplitFile = IIf(IsDBNull(.Fields("SplitFile").Value), False, .Fields("SplitFile").Value And mlngSplitFileSize > 0)

            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mstrExportFileCode = IIf(IsDBNull(.Fields("CMGExportFileCode").Value), "", .Fields("CMGExportFileCode").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mlngCMGRecordIdentifier = IIf(IsDBNull(.Fields("CMGExportRecordID").Value), 0, .Fields("CMGExportRecordID").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mbUpdateAuditLog = IIf(IsDBNull(.Fields("CMGExportUpdateAudit").Value), False, .Fields("CMGExportUpdateAudit").Value)

            mlngExportParent1AllRecords = .Fields("parent1AllRecords").Value
            mlngExportParent1PickListID = .Fields("parent1picklist").Value
            mlngExportParent2AllRecords = .Fields("parent2AllRecords").Value
            mlngExportParent2PickListID = .Fields("parent2picklist").Value

            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mstrExportDelimiter = IIf(IsDBNull(.Fields("delimiter").Value), "", .Fields("delimiter").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mstrExportOtherDelimiter = IIf(IsDBNull(.Fields("otherdelimiter").Value), "", .Fields("otherdelimiter").Value)

            Select Case UCase(mstrExportDelimiter)
                Case "," : mstrExportActualDelimiter = mstrExportDelimiter
                Case "<TAB>" : mstrExportActualDelimiter = vbTab
                Case "<OTHER>" : mstrExportActualDelimiter = mstrExportOtherDelimiter
                Case Else : mstrExportActualDelimiter = mstrExportDelimiter
            End Select

            'mbAppendToFile = IIf(IsNull(!AppendToFile), False, !AppendToFile)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mbForceHeader = IIf(IsDBNull(.Fields("ForceHeader").Value), False, .Fields("ForceHeader").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mbOmitHeader = IIf(IsDBNull(.Fields("OmitHeader").Value), False, .Fields("OmitHeader").Value)

            mlngOutputFormat = .Fields("OutputFormat").Value
            mblnOutputSave = .Fields("OutputSave").Value
            mlngOutputSaveExisting = .Fields("OutputSaveExisting").Value
            'mlngOutputSaveFormat = !OutputSaveFormat
            mblnOutputEmail = .Fields("OutputEmail").Value
            mlngOutputEmailAddr = .Fields("OutputEmailAddr").Value
            mstrOutputEmailSubject = .Fields("OutputEmailSubject").Value
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mstrOutputEmailAttachAs = IIf(IsDBNull(.Fields("OutputEmailAttachAs").Value), vbNullString, .Fields("OutputEmailAttachAs").Value)
            'mlngOutputEmailFileFormat = !OutputEmailFileFormat
            mstrOutputFileName = .Fields("OutputFilename").Value
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mdLastSuccessfulOutput = IIf(IsDBNull(.Fields("LastSuccessfulOutput").Value), "00:00:00", .Fields("LastSuccessfulOutput").Value)

            mintExportHeader = .Fields("Header").Value
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mstrExportHeaderText = IIf(IsDBNull(.Fields("HeaderText").Value), vbNullString, .Fields("HeaderText").Value)
            mintExportFooter = .Fields("Footer").Value
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mstrExportFooterText = IIf(IsDBNull(.Fields("FooterText").Value), vbNullString, .Fields("FooterText").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mstrXMLDataNodeName = IIf(IsDBNull(.Fields("XMLDataNodeName").Value), vbNullString, .Fields("XMLDataNodeName").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mbSplitXMLIntoFiles = IIf(IsDBNull(.Fields("SplitXMLNodesFile").Value), False, .Fields("SplitXMLNodesFile").Value)

            If mlngOutputFormat = modEnums.OutputFormats.fmtXML Then
                mstrExportHeaderText = IIf(Len(mstrExportHeaderText) = 0, "root", mstrExportHeaderText)
                mstrXMLDataNodeName = IIf(Len(mstrXMLDataNodeName) = 0, mstrExportBaseTableName, mstrXMLDataNodeName)
            End If

            '    'If mstrExportOutputType = "S" Then
            '    If mlngOutputFormat = fmtSQLTable Then
            '      Set rsTemp = datGeneral.GetReadOnlyRecords("SELECT TableName FROM ASRSysTables WHERE TableName = '" & mstrOutputFilename & "'")
            '      If Not (rsTemp.BOF And rsTemp.EOF) Then
            '        mstrErrorString = "A table called '" & mstrOutputFilename & "' already exists."
            '        GetExportDefinition = False
            '        Exit Function
            '      End If
            '    End If
            If prstTemp_Definition.Fields("OutputFormat").Value = modEnums.OutputFormats.fmtSQLTable Then
                GetExportDefinition = False
                mstrErrorString = "This Export definition is invalid as export to SQL Table is no longer supported."
                prstTemp_Definition.Close()
                'UPGRADE_NOTE: Object prstTemp_Definition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                prstTemp_Definition = Nothing
                Exit Function
            End If

            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mstrTransformFile = IIf(IsDBNull(.Fields("TransformFile").Value), "", .Fields("TransformFile").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mstrXSDFilename = IIf(IsDBNull(.Fields("XSDFilename").Value), "", .Fields("XSDFilename").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mbPreserveTransformPath = IIf(IsDBNull(.Fields("PreserveTransformPath").Value), False, .Fields("PreserveTransformPath").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mbPreserveXSDPath = IIf(IsDBNull(.Fields("PreserveXSDPath").Value), False, .Fields("PreserveXSDPath").Value)

            mbDefinitionOwner = (LCase(Trim(gsUserName)) = LCase(Trim(.Fields("userName").Value)))
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            mbAuditChangesOnly = IIf(IsDBNull(.Fields("AuditChangesOnly").Value), False, .Fields("AuditChangesOnly").Value)

            If Not IsRecordSelectionValid() Then
                GetExportDefinition = False
                Exit Function
            End If

            If LCase(.Fields("dateseparator").Value) = "<none>" Then
                mstrExportDateSeparator = ""
            Else
                mstrExportDateSeparator = .Fields("dateseparator").Value
            End If

            Select Case .Fields("DateFormat").Value
                Case "dmy"
                    mstrExportDateFormat = "dd" & mstrExportDateSeparator & "mm" & mstrExportDateSeparator & IIf(.Fields("Dateyeardigits").Value = "2", "yy", "yyyy")
                Case "mdy"
                    mstrExportDateFormat = "mm" & mstrExportDateSeparator & "dd" & mstrExportDateSeparator & IIf(.Fields("Dateyeardigits").Value = "2", "yy", "yyyy")
                Case "ymd"
                    mstrExportDateFormat = IIf(.Fields("Dateyeardigits").Value = "2", "yy", "yyyy") & mstrExportDateSeparator & "mm" & mstrExportDateSeparator & "dd"
                Case "ydm"
                    mstrExportDateFormat = IIf(.Fields("Dateyeardigits").Value = "2", "yy", "yyyy") & mstrExportDateSeparator & "dd" & mstrExportDateSeparator & "mm"
            End Select

        End With


        If mlngOutputFormat = modEnums.OutputFormats.fmtExcelWorksheet And Not gblnBatchMode Then
            gobjProgress.AVI = COAProgress.AVIType.dbExcel
        End If


        GetExportDefinition = True

TidyAndExit:

        'UPGRADE_NOTE: Object prstTemp_Definition may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        prstTemp_Definition = Nothing

        Exit Function

GetExportDefinition_ERROR:

        GetExportDefinition = False
        mstrErrorString = "Error whilst reading the Export definition !" & vbCrLf & "(" & Err.Description & ")"
        Resume TidyAndExit

    End Function

    Private Function GetDetailsRecordsets() As Boolean

        ' Purpose : This function loads export details and sort details into
        '           arrays and leaves the details recordset reference there
        '           (dont remove it...used for summary info !) NB CAN REMOVE FOR EXPORT !?!
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo GetDetailsRecordsets_ERROR

        Dim pstrTempSQL As String
        Dim pintTemp As Short
        Dim prstExportSortOrder As ADODB.Recordset

        Dim lngTextCount As Integer
        Dim lngFillerCount As Integer
        Dim lngRecNumCount As Integer
        Dim lngCalculationCount As Integer

        'gobjProgress.UpdateProgress gblnBatchMode

        ' If user cancels the export, abort
        If gobjProgress.Cancelled Then
            mblnUserCancelled = True
            GetDetailsRecordsets = False
            Exit Function
        End If

        ' Get the column information from the Details table, in order

        pstrTempSQL = "SELECT * FROM AsrSysExportDetails WHERE " & "ExportID = " & mlngExportID & " " & "ORDER BY [ID]"

        mrstExportDetails = mclsData.OpenRecordset(pstrTempSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        With mrstExportDetails

            If .BOF And .EOF Then
                GetDetailsRecordsets = False
                mstrErrorString = "No columns found in the specified Export definition." & vbCrLf & "Please remove this definition and create a new one."
                Exit Function
            End If

            Do Until .EOF
                pintTemp = UBound(mvarColDetails, 2) + 1
                ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), pintTemp)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(0, pintTemp) = .Fields("Type").Value
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(1, pintTemp) = Trim(.Fields("TableID").Value)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(2, pintTemp) = datGeneral.GetTableName(CInt(.Fields("TableID").Value))
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(3, pintTemp) = .Fields("ColExprID").Value
                If .Fields("ColExprID").Value = 0 Or .Fields("Type").Value = "X" Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mvarColDetails(4, pintTemp) = ""
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mvarColDetails(4, pintTemp) = datGeneral.GetColumnName(CInt(.Fields("ColExprID").Value))
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(5, pintTemp) = .Fields("Data").Value
                'mvarColDetails(6, pintTemp) = !fillerlength
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(6, pintTemp) = IIf(.Fields("fillerlength").Value > 999999, 999999, .Fields("fillerlength").Value)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(7, pintTemp) = datGeneral.DateColumn(.Fields("Type").Value, .Fields("TableID").Value, .Fields("ColExprID").Value)

                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(8, pintTemp) = IIf(IsDBNull(.Fields("CMGColumnCode").Value), "", .Fields("CMGColumnCode").Value) ' The CMG Code for this column
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(9, pintTemp) = datGeneral.IsColumnAudited(.Fields("ColExprID").Value) ' Is this column audited
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(10, pintTemp) = True ' Do we export this column

                'TM20011010 Fault 2197
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(11, pintTemp) = .Fields("Decimals").Value

                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(12, pintTemp) = datGeneral.NumericColumn(.Fields("Type").Value, .Fields("TableID").Value, .Fields("ColExprID").Value)


                'Default to old heading
                '(still need to increase counter even if heading is overwritten!)
                Select Case .Fields("Type").Value
                    Case "F", "R" 'Filler or Carriage Return
                        lngFillerCount = lngFillerCount + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mvarColDetails(13, pintTemp) = "Filler" & CStr(lngFillerCount)
                    Case "T" 'Text
                        lngTextCount = lngTextCount + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mvarColDetails(13, pintTemp) = "Text" & CStr(lngTextCount)
                    Case "N" 'Record Number
                        lngRecNumCount = lngRecNumCount + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mvarColDetails(13, pintTemp) = "Record Number" & CStr(lngRecNumCount)
                    Case "X" 'Calc
                        lngCalculationCount = lngCalculationCount + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mvarColDetails(13, pintTemp) = "Calculation" & CStr(lngCalculationCount)
                End Select

                'MH20030120
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                If Not IsDBNull(.Fields("Heading").Value) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mvarColDetails(13, pintTemp) = .Fields("Heading").Value 'might be null but thats okay!
                End If
                'UPGRADE_WARNING: Couldn't resolve default property of object GetUniqueHeading(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(14, pintTemp) = Replace(GetUniqueHeading(pintTemp), "'", "''")

                ' Does this column use digit seperators
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(15, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(15, pintTemp) = datGeneral.DoesColumnUseSeparators(.Fields("ColExprID").Value)

                'NPG20071217 Fault 12867
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(16, pintTemp) = .Fields("ConvertCase").Value

                'NPG20080617 Suggestion S000816
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(17, pintTemp) = .Fields("SuppressNulls").Value

                .MoveNext()
            Loop
            .MoveFirst()
        End With

        'Add in the ID of the record if we are a cmg output
        'If mstrExportOutputType = "C" Then
        If mlngOutputFormat = modEnums.OutputFormats.fmtCMGFile Or mlngOutputFormat = modEnums.OutputFormats.fmtXML Then
            pintTemp = UBound(mvarColDetails, 2) + 1
            ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), pintTemp)
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(0, pintTemp) = "C"
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(1, pintTemp) = mlngExportBaseTable
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(2, pintTemp) = mstrExportBaseTableName
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(3, pintTemp) = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(4, pintTemp) = "ID"
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(5, pintTemp) = "ID"
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(6, pintTemp) = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(7, pintTemp) = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(8, pintTemp) = "" ' The CMG Code for this column
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(9, pintTemp) = False ' Is this column audited
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(10, pintTemp) = False ' Do we export this column
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(11, pintTemp) = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(12, pintTemp) = False
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(13, pintTemp) = "ID"
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(14, pintTemp) = "ID"
        End If

        If mlngOutputFormat = modEnums.OutputFormats.fmtCMGFile Then
            pintTemp = UBound(mvarColDetails, 2) + 1
            ReDim Preserve mvarColDetails(UBound(mvarColDetails, 1), pintTemp)
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(0, pintTemp) = "C"
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(1, pintTemp) = mlngExportBaseTable
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(2, pintTemp) = mstrExportBaseTableName
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(3, pintTemp) = mlngCMGRecordIdentifier
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(4, pintTemp) = datGeneral.GetColumnName(mlngCMGRecordIdentifier)
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(5, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(5, pintTemp) = "Identifier"
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(6, pintTemp) = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(7, pintTemp) = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(8, pintTemp) = "" ' The CMG Code for this column
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(9, pintTemp) = False ' Is this column audited
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(10, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(10, pintTemp) = False ' Do we export this column
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(11, pintTemp) = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(12, pintTemp) = False
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(13, pintTemp) = "Identifier"
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mvarColDetails(14, pintTemp) = "Identifier"

        End If

        ' Get those columns defined as a SortOrder and load into array

        pstrTempSQL = "SELECT * FROM AsrSysExportDetails WHERE " & "ExportID = " & mlngExportID & " " & "AND SortOrderSequence > 0 AND Type = 'C' " & "ORDER BY [SortOrderSequence]"

        prstExportSortOrder = datGeneral.GetReadOnlyRecords(pstrTempSQL)

        With prstExportSortOrder
            If .BOF And .EOF Then
                GetDetailsRecordsets = False
                mstrErrorString = "No columns have been defined as a sort order for the specified Export definition." & vbCrLf & "Please remove this definition and create a new one."
                Exit Function
            End If
            Do Until .EOF
                pintTemp = UBound(mvarSortOrder, 2) + 1
                ReDim Preserve mvarSortOrder(2, pintTemp)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarSortOrder(1, pintTemp) = .Fields("ColExprID").Value
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, pintTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarSortOrder(2, pintTemp) = .Fields("SortOrder").Value
                .MoveNext()
            Loop
        End With

        'UPGRADE_NOTE: Object prstExportSortOrder may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        prstExportSortOrder = Nothing
        'UPGRADE_NOTE: Object mrstExportDetails may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mrstExportDetails = Nothing

        GetDetailsRecordsets = True
        Exit Function

GetDetailsRecordsets_ERROR:

        GetDetailsRecordsets = False
        mstrErrorString = "Error whilst retrieving the details recordsets." & vbCrLf & "(" & Err.Description & ")"

    End Function

    Private Function GenerateSQL() As Boolean

        ' Purpose : This function calls the individual functions that
        '           general the components of the main SQL string.
        ' Input   : None
        ' Output  : True/False Success

        Dim fOK As Boolean

        'gobjProgress.UpdateProgress gblnBatchMode

        ' If user cancels the export, abort
        If gobjProgress.Cancelled Then
            mblnUserCancelled = True
            GenerateSQL = False
            Exit Function
        End If

        fOK = True

        If fOK Then fOK = GenerateSQLSelect()
        If fOK Then fOK = GenerateSQLFrom()
        If fOK Then fOK = GenerateSQLJoin()
        If fOK Then fOK = GenerateSQLWhere()
        If fOK Then fOK = GenerateSQLOrderBy()

        If fOK Then GenerateSQL = True Else GenerateSQL = False

    End Function

    Private Function GenerateSQLSelect() As Boolean

        ' Purpose : This function compiles the SQLSelect string looping
        '           thru the column details recordset.
        '           NB. NEEDS TO BE CHANGED FOR EXPRESSIONS !!!
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo GenerateSQLSelect_ERROR

        Dim plngTempTableID As Integer
        Dim pstrTempTableName As String
        Dim pstrTempColumnName As String

        Dim pblnOK As Boolean
        Dim pblnColumnOK As Boolean

        Dim pblnNoSelect As Boolean
        Dim pblnFound As Boolean

        Dim pintLoop As Short
        Dim pstrColumnList As String
        Dim pstrColumnCode As String
        Dim pstrSource As String
        Dim pintNextIndex As Short

        Dim alngSourceTables(,) As Integer
        Dim sCalcCode As String
        Dim blnOK As Boolean
        Dim objCalcExpr As New clsExprExpression

        Dim iTextCount As Short
        Dim iFillerCount As Short
        Dim iRecNumCount As Short
        Dim iCalculationCount As Short
        Dim iLoop1 As Short

        Dim sCalcName As String
        Dim objTableView As TablePrivilege

        ' Set flags with their starting values
        pblnOK = True
        pblnNoSelect = False
        pstrColumnList = ""

        ReDim mastrUDFsRequired(0)

        ' JPD20030219 Fault 5066
        ' Check the user has permission to read the base table.
        pblnOK = False
        For Each objTableView In gcoTablePrivileges.Collection
            If (objTableView.TableID = mlngExportBaseTable) And (objTableView.AllowSelect) Then
                pblnOK = True
                Exit For
            End If
        Next objTableView
        'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objTableView = Nothing

        If Not pblnOK Then
            GenerateSQLSelect = False
            mstrErrorString = "You do not have permission to read the base table" & vbCrLf & "either directly or through any views."
            Exit Function
        End If

        ' COWBOY ALERT !!!!!!! (Forgive me)
        ' JDM - 05/09/2005 - Fault 1030X - SQL 2005 ORDER BY clause does not work if used in conjunction with the INTO clause.
        '                                  This looks like its by design and not just a beta fault, but we can re-invetiagte when
        '                                  Micr*s*ft release the full product.
        ' Start off the select statement

        ' COWBOY ALERT REVISITED !!!!!!!
        ' JDM - 20/7/2012 - JIRA xxxx - Congratulations to Microsoft - you have now reintroduced this with SQL 2012. Exactly the same issue
        '                               as with SQL 2005, for some reason when using the SELECT... INTO it ignores the sort.
        '                               Run the query analyser with display execution plan to find out for yourself.
        '                               This doesn't affect 2008. I wonder what the next version of SQL will do?

        ' COWBOY ALERT RE-REVISTITED !!!!!!
        ' JDM - 17/09/2014 - TFS-9973 - Well, well, well, here we go again.
        If glngSQLVersion = 9 Or glngSQLVersion >= 10.5 Then
            mstrSQLSelect = "SELECT TOP 1000000000000 "
        Else
            mstrSQLSelect = "SELECT "
        End If

        ' Dimension an array of tables/views joined to the base table/view
        ' Column 1 = 0 if this row is for a table, 1 if it is for a view
        ' Column 2 = table/view ID
        ' (should contain everything which needs to be joined to the base tbl/view)
        ReDim mlngTableViews(2, 0)

        ' Loop thru the columns collection creating the SELECT and JOIN code
        iRecNumCount = 0
        For pintLoop = 1 To UBound(mvarColDetails, 2)

            ' Clear temp vars
            plngTempTableID = 0
            pstrTempTableName = vbNullString
            pstrTempColumnName = vbNullString

            'C - Column
            'T - Text
            'F - Filler
            'X - Expression
            'Y - Carriage return


            ' If its a COLUMN then...
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If mvarColDetails(0, pintLoop) = "C" Then

                ' Load the temp variables
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                plngTempTableID = mvarColDetails(1, pintLoop)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(2, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pstrTempTableName = mvarColDetails(2, pintLoop)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pstrTempColumnName = mvarColDetails(4, pintLoop)

                ' Check permission on that column
                mobjColumnPrivileges = GetColumnPrivileges(pstrTempTableName)
                mstrRealSource = gcoTablePrivileges.item(pstrTempTableName).RealSource

                pblnColumnOK = mobjColumnPrivileges.IsValid(pstrTempColumnName)

                If pblnColumnOK Then
                    pblnColumnOK = mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect
                End If

                If pblnColumnOK Then

                    ' this column can be read direct from the tbl/view or from a parent table

                    '        pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & _
                    ''        mstrRealSource & "." & Trim(pstrTempColumnName) & _
                    ''        " AS '" & mvarColDetails(14, pintLoop) & "'"
                    '        '" AS '" & mvarColDetails(5, pintLoop) & "'"


                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & mstrRealSource & "." & Trim(pstrTempColumnName) & " AS '" & mvarColDetails(14, pintLoop) & "'"


                    ' If the table isnt the base table (or its realsource) then
                    ' Check if it has already been added to the array. If not, add it.
                    If plngTempTableID <> mlngExportBaseTable Then
                        pblnFound = False
                        For pintNextIndex = 1 To UBound(mlngTableViews, 2)
                            If mlngTableViews(1, pintNextIndex) = 0 And mlngTableViews(2, pintNextIndex) = plngTempTableID Then
                                pblnFound = True
                                Exit For
                            End If
                        Next pintNextIndex

                        If Not pblnFound Then
                            pintNextIndex = UBound(mlngTableViews, 2) + 1
                            ReDim Preserve mlngTableViews(2, pintNextIndex)
                            mlngTableViews(1, pintNextIndex) = 0
                            mlngTableViews(2, pintNextIndex) = plngTempTableID
                        End If
                    End If

                    ' Optional filter for only auditted changes
                    If mdLastSuccessfulOutput <> CDate("00:00:00") Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mstrOnlyChangesFilter = CDbl(mstrOnlyChangesFilter) + IIf(Len(mstrOnlyChangesFilter) > 0, " OR ", "") & " dbo.udfsys_FieldChangedSinceLastExport(" & mvarColDetails(3, pintLoop) & ", '" & VB6.Format(mdLastSuccessfulOutput, "MM/dd/yyyy hh:mm:ss") & "', " & mstrRealSource & ".[id]) = 1"
                    End If

                Else

                    ' this column cannot be read direct. If its from a parent, try parent views
                    ' Loop thru the views on the table, seeing if any have read permis for the column

                    ReDim mstrViews(0)
                    For Each mobjTableView In gcoTablePrivileges.Collection
                        If (Not mobjTableView.IsTable) And (mobjTableView.TableID = plngTempTableID) And (mobjTableView.AllowSelect) Then

                            pstrSource = mobjTableView.ViewName
                            mstrRealSource = gcoTablePrivileges.item(pstrSource).RealSource

                            ' Get the column permission for the view
                            mobjColumnPrivileges = GetColumnPrivileges(pstrSource)

                            ' If we can see the column from this view
                            If mobjColumnPrivileges.IsValid(pstrTempColumnName) Then
                                If mobjColumnPrivileges.Item(pstrTempColumnName).AllowSelect Then

                                    ReDim Preserve mstrViews(UBound(mstrViews) + 1)
                                    mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

                                    ' Check if view has already been added to the array
                                    pblnFound = False
                                    For pintNextIndex = 1 To UBound(mlngTableViews, 2)
                                        If mlngTableViews(1, pintNextIndex) = 1 And mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID Then
                                            pblnFound = True
                                            Exit For
                                        End If
                                    Next pintNextIndex

                                    If Not pblnFound Then

                                        ' View hasnt yet been added, so add it !
                                        pintNextIndex = UBound(mlngTableViews, 2) + 1
                                        ReDim Preserve mlngTableViews(2, pintNextIndex)
                                        mlngTableViews(1, pintNextIndex) = 1
                                        mlngTableViews(2, pintNextIndex) = mobjTableView.ViewID

                                    End If
                                End If
                            End If
                        End If

                    Next mobjTableView

                    'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                    mobjTableView = Nothing

                    ' Does the user have select permission thru ANY views ?
                    If UBound(mstrViews) = 0 Then
                        pblnNoSelect = True
                    Else

                        ' Add the column to the column list
                        pstrColumnCode = ""
                        For pintNextIndex = 1 To UBound(mstrViews)
                            If pintNextIndex = 1 Then
                                pstrColumnCode = "CASE"
                            End If

                            pstrColumnCode = pstrColumnCode & " WHEN NOT " & mstrViews(pintNextIndex) & "." & pstrTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & pstrTempColumnName

                        Next pintNextIndex

                        If Len(pstrColumnCode) > 0 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            pstrColumnCode = pstrColumnCode & " ELSE NULL" & " END AS '" & mvarColDetails(14, pintLoop) & "'"
                            '" END AS '" & mvarColDetails(5, pintLoop) & "'"


                            pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & pstrColumnCode
                        End If

                    End If

                    ' If we cant see a column, then get outta here
                    If pblnNoSelect Then
                        GenerateSQLSelect = False
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mstrErrorString = "You do not have permission to see the column '" & mvarColDetails(4, pintLoop) & "'" & vbCrLf & "either directly or through any views." & vbCrLf & vbCrLf & "You may not run this Export."
                        Exit Function
                    End If


                    If Not pblnOK Then
                        GenerateSQLSelect = False
                        Exit Function
                    End If

                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf mvarColDetails(0, pintLoop) = "T" Then

                ' Text, increase the text counter...
                iTextCount = iTextCount + 1
                'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
                '"'" & Replace(mvarColDetails(5, pintLoop), "'", "''") & "' AS Text" & iTextCount
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & "'" & Replace(mvarColDetails(5, pintLoop), "'", "''") & "' AS '" & mvarColDetails(14, pintLoop) & "'"

                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf mvarColDetails(0, pintLoop) = "R" Then

                'Need to allow two characters for carriage return and line feed
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarColDetails(6, pintLoop) = 2

                ' Do the same as filler, increase the filer counter.
                iFillerCount = iFillerCount + 1
                'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
                '"'" & vbCrLf & "' AS Filler" & iFillerCount
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & "'" & vbCrLf & "' AS '" & mvarColDetails(14, pintLoop) & "'"

                '"0 AS Filler" & iFillerCount

                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf mvarColDetails(0, pintLoop) = "N" Then

                ' Do the same as filler, increase the filer counter.
                iRecNumCount = iRecNumCount + 1
                'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
                '"'" & vbCrLf & "' AS 'Record Number" & IIf(iRecNumCount > 1, CStr(iRecNumCount), "") & "'"
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & "'" & vbCrLf & "' AS '" & mvarColDetails(14, pintLoop) & "'"

                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf mvarColDetails(0, pintLoop) = "F" Then

                ' Filler, increase the filer counter.
                iFillerCount = iFillerCount + 1

                ' If export is fixed length, use filler length as defined else use 1.
                'If mstrExportOutputType <> "F" Then
                If mlngOutputFormat <> modEnums.OutputFormats.fmtFixedLengthFile Then
                    'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
                    '"' ' AS Filler" & iFillerCount
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & "' ' AS '" & mvarColDetails(14, pintLoop) & "'"
                Else
                    'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
                    '"'" & Space(mvarColDetails(6, pintLoop)) & "' AS Filler" & iFillerCount
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & "'" & Space(mvarColDetails(6, pintLoop)) & "' AS '" & mvarColDetails(14, pintLoop) & "'"
                End If

                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ElseIf mvarColDetails(0, pintLoop) = "X" Then

                ' Its an expression !!!

                iCalculationCount = iCalculationCount + 1

                '      pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ", ", "") & _
                '"'" & Replace(mvarColDetails(5, pintLoop), "'", "''") & "' AS Calculation" & iCalculationCount


                '################ NICKED FROM CUSTOM REPORT CODE....
                '
                '      ' Get the calculation SQL, and the array of tables/views that are used to create it.
                '      ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
                '      ' Column 2 = table/view ID.
                ReDim alngSourceTables(2, 0)
                objCalcExpr = New clsExprExpression
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                blnOK = objCalcExpr.Initialise(mlngExportBaseTable, CInt(mvarColDetails(3, pintLoop)), modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION, modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
                sCalcName = objCalcExpr.Name
                If blnOK Then
                    blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, True)

                    If blnOK And gbEnableUDFFunctions Then
                        blnOK = objCalcExpr.UDFCalculationCode(alngSourceTables, mastrUDFsRequired, True)
                    End If

                End If

                'TM20030422 Fault 5242 - The "SELECT ... INTO..." statement errors when it trys to create a column for
                'and empty string. Therefore wrap this empty sting in a CONVERT(varchar... clause.
                'TM20030521 Fault 5702 - Compare the empty string with the calc code value converted to varchar
                sCalcCode = "CASE WHEN CONVERT(varchar," & sCalcCode & ") = '' " & "THEN CONVERT(varchar," & sCalcCode & ") " & "ELSE " & sCalcCode & " END"

                'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objCalcExpr = Nothing

                If blnOK Then

                    'pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & _
                    'sCalcCode & " AS 'Calculation" & iCalculationCount & "' "
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    pstrColumnList = pstrColumnList & IIf(Len(pstrColumnList) > 0, ",", "") & sCalcCode & " AS '" & mvarColDetails(14, pintLoop) & "' "

                    ' Add the required views to the JOIN code.
                    For iLoop1 = 1 To UBound(alngSourceTables, 2)
                        If alngSourceTables(1, iLoop1) = 1 Then
                            ' Check if view has already been added to the array
                            pblnFound = False
                            For pintNextIndex = 1 To UBound(mlngTableViews, 2)
                                If mlngTableViews(1, pintNextIndex) = 1 And mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1) Then
                                    pblnFound = True
                                    Exit For
                                End If
                            Next pintNextIndex

                            If Not pblnFound Then

                                ' View hasnt yet been added, so add it !
                                pintNextIndex = UBound(mlngTableViews, 2) + 1
                                ReDim Preserve mlngTableViews(2, pintNextIndex)
                                mlngTableViews(1, pintNextIndex) = 1
                                mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1)

                            End If

                            'TM20020121 Fault 2277
                            '********************************************************************************
                        ElseIf alngSourceTables(1, iLoop1) = 0 Then
                            ' Check if table has already been added to the array
                            pblnFound = False
                            For pintNextIndex = 1 To UBound(mlngTableViews, 2)
                                If mlngTableViews(1, pintNextIndex) = 0 And mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1) Then
                                    pblnFound = True
                                    Exit For
                                End If
                            Next pintNextIndex

                            If Not pblnFound Then
                                ' table hasnt yet been added, so add it !
                                pintNextIndex = UBound(mlngTableViews, 2) + 1
                                ReDim Preserve mlngTableViews(2, pintNextIndex)
                                mlngTableViews(1, pintNextIndex) = 0
                                mlngTableViews(2, pintNextIndex) = alngSourceTables(2, iLoop1)
                            End If
                            '********************************************************************************
                        End If
                    Next iLoop1
                Else
                    ' Permission denied on something in the calculation.
                    mstrErrorString = "You do not have permission to use the '" & sCalcName & "' calculation."
                    GenerateSQLSelect = False
                    Exit Function
                End If

                '################
            End If

        Next pintLoop

        mstrSQLSelect = mstrSQLSelect & pstrColumnList

        If mlngOutputFormat = modEnums.OutputFormats.fmtXML Then
            mstrSQLSelect = mstrSQLSelect & ", NEWID() AS [_XMLSplitID]"
        End If

        GenerateSQLSelect = True

        Exit Function

GenerateSQLSelect_ERROR:

        GenerateSQLSelect = False
        mstrErrorString = "Error whilst generating SQL Select statement." & vbCrLf & "(" & Err.Description & ")"

    End Function

    Private Function GenerateSQLFrom() As Boolean

        ' Purpose : It doesnt take Einstein to work out that this function
        '           adds the base table name to the from clause of the SQL string.
        ' Input   : None
        ' Output  : True/False Success

        '  Dim iLoop As Integer
        Dim pobjTableView As TablePrivilege

        pobjTableView = New TablePrivilege

        mstrSQLFrom = gcoTablePrivileges.item(mstrExportBaseTableName).RealSource

        '  Else
        '
        '    ' need some way of determining which view out of all these
        '    ' is based on the base table, use it in the from statement
        '    ' and mark it so that i know what tables to join to it
        '
        '    ' NB 13/03/00 IS THIS NEEDED NOW AS ALL USERS WILL HAVE SELECT PERMISSION
        '    '             ON THE ID COLUMN OF ALL TABLES ???
        '
        '    For Each pobjTableView In gcoTablePrivileges.Collection
        '
        '      If (Not pobjTableView.IsTable) And _
        ''      (pobjTableView.TableID = mlngExportBaseTable) And _
        ''      (pobjTableView.AllowSelect) Then
        '
        '        mstrSQLFrom = pobjTableView.ViewName
        '        Exit For
        '
        '    End If
        '
        '    Next pobjTableView
        '
        '  End If

        'UPGRADE_NOTE: Object pobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        pobjTableView = Nothing

        GenerateSQLFrom = True
        Exit Function

GenerateSQLFrom_ERROR:

        GenerateSQLFrom = False
        mstrErrorString = "Error in GenerateSQLFrom." & vbCrLf & "(" & Err.Description & ""

    End Function

    Private Function GenerateSQLJoin() As Boolean

        ' Purpose : Add the join strings for parent/child/views.
        '           Also adds filter clauses to the joins if used
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo GenerateSQLJoin_ERROR

        Dim blnOK As Boolean
        Dim pobjTableView As TablePrivilege
        Dim objChildTable As TablePrivilege
        Dim pintLoop As Short
        Dim sChildJoinCode As String
        Dim sReuseJoinCode As String
        Dim sChildOrderString As String
        Dim rsTemp As ADODB.Recordset
        Dim strFilterIDs As String
        Dim pblnChildUsed As Boolean
        Dim sChildJoin As String
        Dim objExpr As clsExprExpression

        ' Get the base table real source
        mstrBaseTableRealSource = mstrSQLFrom

        ' First, do the joins for all the views etc...
        For pintLoop = 1 To UBound(mlngTableViews, 2)

            ' Get the table/view object from the id stored in the array
            If mlngTableViews(1, pintLoop) = 0 Then
                pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
            Else
                pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
            End If

            ' Dont add a join here if its the child or parent table...do that later
            If pobjTableView.TableID <> mlngExportChildTable Then
                If pobjTableView.TableID <> mlngExportParent1Table Then
                    If pobjTableView.TableID <> mlngExportParent2Table Then

                        If (pobjTableView.ViewName <> mstrBaseTableRealSource) Then
                            mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & pobjTableView.RealSource & ".ID"
                        End If

                    End If
                End If
            End If

            ' OK, parent table joins...
            If (pobjTableView.TableID = mlngExportParent1Table) Or (pobjTableView.TableID = mlngExportParent2Table) Then

                mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & pobjTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID_" & pobjTableView.TableID & " = " & pobjTableView.RealSource & ".ID"
            End If

        Next pintLoop

        ' Second, do the childview bit, if required

        If mlngExportChildTable > 0 Then

            ' are any child fields in the export ? # 12/06/00 RH - FAULT 419
            For pintLoop = 1 To UBound(mvarColDetails, 2)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(1, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If mvarColDetails(1, pintLoop) = mlngExportChildTable Then pblnChildUsed = True
            Next pintLoop

            If pblnChildUsed = True Then

                objChildTable = gcoTablePrivileges.FindTableID(mlngExportChildTable)

                If objChildTable.AllowSelect Then
                    sChildJoinCode = " LEFT OUTER JOIN " & objChildTable.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & objChildTable.RealSource & ".ID_" & mlngExportBaseTable

                    sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN"

                    sChildJoinCode = sChildJoinCode & " (SELECT TOP" & IIf(mlngExportChildMaxRecords = 0, " 100 PERCENT", " " & mlngExportChildMaxRecords) & " " & objChildTable.RealSource & ".ID FROM " & objChildTable.RealSource

                    ' Now the child order by bit - done here in case tables need to be joined.
                    rsTemp = datGeneral.GetOrderDefinition(datGeneral.GetDefaultOrder(mlngExportChildTable))
                    sChildOrderString = DoChildOrderString(rsTemp, sChildJoin)
                    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                    rsTemp = Nothing

                    sChildJoinCode = sChildJoinCode & sChildJoin

                    sChildJoinCode = sChildJoinCode & " WHERE (" & objChildTable.RealSource & ".ID_" & mlngExportBaseTable & " = " & mstrBaseTableRealSource & ".ID)"

                    ' is the child filtered ?

                    If mlngExportChildFilterID > 0 Then
                        blnOK = datGeneral.FilteredIDs(mlngExportChildFilterID, strFilterIDs)

                        ' Generate any UDFs that are used in this filter
                        If blnOK Then
                            datGeneral.FilterUDFs(mlngExportChildFilterID, mastrUDFsRequired)
                        End If

                        If blnOK Then
                            sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN (" & strFilterIDs & ")"
                        Else
                            ' Permission denied on something in the filter.
                            mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngExportChildFilterID) & "' filter."
                            GenerateSQLJoin = False
                            Exit Function
                        End If
                    End If

                End If

            End If

        End If

        mstrSQLJoin = mstrSQLJoin & sChildJoinCode & IIf(Len(sChildOrderString) > 0, " ORDER BY " & sChildOrderString & ")", "")

        GenerateSQLJoin = True
        Exit Function

GenerateSQLJoin_ERROR:

        GenerateSQLJoin = False
        mstrErrorString = "Error in GenerateSQLJoin." & vbCrLf & Err.Description

    End Function

    Private Function DoChildOrderString(ByRef rsTemp As ADODB.Recordset, ByRef psJoinCode As String) As String

        ' This function loops through the child tables default order
        ' checking if the user has privileges. If they do, add to the order string
        ' if not, leave it out.

        On Error GoTo DoChildOrderString_ERROR

        Dim fColumnOK As Boolean
        Dim fFound As Boolean
        Dim iNextIndex As Short
        Dim sSource As String
        Dim sRealSource As String
        Dim sColumnCode As String
        Dim sCurrentTableViewName As String
        Dim objColumnPrivileges As CColumnPrivileges
        Dim pobjOrderCol As TablePrivilege
        Dim objTableView As TablePrivilege
        Dim alngTableViews(,) As Integer
        Dim asViews() As String
        Dim iTempCounter As Short

        ' Dimension an array of tables/views joined to the base table/view.
        ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
        ' Column 2 = table/view ID.
        ReDim alngTableViews(2, 0)

        pobjOrderCol = gcoTablePrivileges.FindTableID(mlngExportChildTable)
        sCurrentTableViewName = pobjOrderCol.RealSource
        'UPGRADE_NOTE: Object pobjOrderCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        pobjOrderCol = Nothing

        Do Until rsTemp.EOF
            If rsTemp.Fields("Type").Value = "O" Then
                ' Check if the user can read the column.
                pobjOrderCol = gcoTablePrivileges.FindTableID(rsTemp.Fields("TableID").Value)
                objColumnPrivileges = GetColumnPrivileges((pobjOrderCol.TableName))
                fColumnOK = objColumnPrivileges.Item(rsTemp.Fields("ColumnName")).AllowSelect
                'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                objColumnPrivileges = Nothing

                If fColumnOK Then
                    If rsTemp.Fields("TableID").Value = mlngExportChildTable Then
                        DoChildOrderString = DoChildOrderString & IIf(Len(DoChildOrderString) > 0, ",", "") & pobjOrderCol.RealSource & "." & rsTemp.Fields("ColumnName").Value & IIf(rsTemp.Fields("Ascending").Value, "", " DESC")
                    Else
                        ' If the column comes from a parent table, then add the table to the Join code.
                        ' Check if the table has already been added to the join code.
                        fFound = False
                        iTempCounter = 0
                        For iNextIndex = 1 To UBound(alngTableViews, 2)
                            If alngTableViews(1, iNextIndex) = 0 And alngTableViews(2, iNextIndex) = rsTemp.Fields("TableID").Value Then
                                iTempCounter = iNextIndex
                                fFound = True
                                Exit For
                            End If
                        Next iNextIndex

                        If Not fFound Then
                            ' The table has not yet been added to the join code, so add it to the array and the join code.
                            iNextIndex = UBound(alngTableViews, 2) + 1
                            ReDim Preserve alngTableViews(2, iNextIndex)
                            alngTableViews(1, iNextIndex) = 0
                            alngTableViews(2, iNextIndex) = rsTemp.Fields("TableID").Value

                            iTempCounter = iNextIndex

                            psJoinCode = psJoinCode & " LEFT OUTER JOIN " & pobjOrderCol.RealSource & " ASRSysTemp_" & Trim(Str(iTempCounter)) & " ON " & sCurrentTableViewName & ".ID_" & Trim(Str(rsTemp.Fields("TableID").Value)) & " = ASRSysTemp_" & Trim(Str(iTempCounter)) & ".ID"
                        End If

                        DoChildOrderString = DoChildOrderString & IIf(Len(DoChildOrderString) > 0, ",", "") & "ASRSysTemp_" & Trim(Str(iTempCounter)) & "." & rsTemp.Fields("ColumnName").Value & IIf(rsTemp.Fields("Ascending").Value, "", " DESC")
                    End If
                Else
                    ' The column cannot be read from the base table/view, or directly from a parent table.
                    ' If it is a column from a prent table, then try to read it from the views on the parent table.
                    If rsTemp.Fields("TableID").Value <> mlngExportChildTable Then
                        ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
                        ReDim asViews(0)
                        For Each objTableView In gcoTablePrivileges.Collection
                            If (Not objTableView.IsTable) And (objTableView.TableID = rsTemp.Fields("TableID").Value) And (objTableView.AllowSelect) Then

                                sSource = objTableView.ViewName
                                sRealSource = gcoTablePrivileges.item(sSource).RealSource

                                ' Get the column permission for the view.
                                objColumnPrivileges = GetColumnPrivileges(sSource)

                                If objColumnPrivileges.IsValid(rsTemp.Fields("ColumnName")) Then
                                    If objColumnPrivileges.Item(rsTemp.Fields("ColumnName")).AllowSelect Then
                                        ' Add the view info to an array to be put into the column list or order code below.
                                        iNextIndex = UBound(asViews) + 1
                                        ReDim Preserve asViews(iNextIndex)
                                        asViews(iNextIndex) = objTableView.ViewName

                                        ' Add the view to the Join code.
                                        ' Check if the view has already been added to the join code.
                                        fFound = False
                                        iTempCounter = 0
                                        For iNextIndex = 1 To UBound(alngTableViews, 2)
                                            If alngTableViews(1, iNextIndex) = 1 And alngTableViews(2, iNextIndex) = objTableView.ViewID Then
                                                fFound = True
                                                iTempCounter = iNextIndex
                                                Exit For
                                            End If
                                        Next iNextIndex

                                        If Not fFound Then
                                            ' The view has not yet been added to the join code, so add it to the array and the join code.
                                            iNextIndex = UBound(alngTableViews, 2) + 1
                                            ReDim Preserve alngTableViews(2, iNextIndex)
                                            alngTableViews(1, iNextIndex) = 1
                                            alngTableViews(2, iNextIndex) = objTableView.ViewID

                                            iTempCounter = iNextIndex

                                            psJoinCode = psJoinCode & " LEFT OUTER JOIN " & sRealSource & " ASRSysTemp_" & Trim(Str(iTempCounter)) & " ON " & sCurrentTableViewName & ".ID_" & Trim(Str(objTableView.TableID)) & " = ASRSysTemp_" & Trim(Str(iTempCounter)) & ".ID"
                                        End If
                                    End If
                                End If
                                'UPGRADE_NOTE: Object objColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                                objColumnPrivileges = Nothing
                            End If
                        Next objTableView
                        'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                        objTableView = Nothing

                        ' The current user does have permission to 'read' the column through a/some view(s) on the
                        ' table.
                        If UBound(asViews) > 0 Then
                            ' Add the column to the column list.
                            sColumnCode = ""
                            For iNextIndex = 1 To UBound(asViews)
                                If iNextIndex = 1 Then
                                    sColumnCode = "CASE "
                                End If

                                sColumnCode = sColumnCode & " WHEN NOT ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & rsTemp.Fields("ColumnName").Value & " IS NULL THEN ASRSysTemp_" & Trim(Str(iNextIndex)) & "." & rsTemp.Fields("ColumnName").Value
                            Next iNextIndex

                            If Len(sColumnCode) > 0 Then
                                sColumnCode = sColumnCode & " ELSE NULL" & " END"

                                ' Add the column to the order string.
                                DoChildOrderString = DoChildOrderString & IIf(Len(DoChildOrderString) > 0, ", ", "") & sColumnCode & IIf(rsTemp.Fields("Ascending").Value, "", " DESC")
                            End If
                        End If
                    End If
                End If

                'UPGRADE_NOTE: Object pobjOrderCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                pobjOrderCol = Nothing
            End If

            rsTemp.MoveNext()
        Loop

        Exit Function

DoChildOrderString_ERROR:

        'UPGRADE_NOTE: Object pobjOrderCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        pobjOrderCol = Nothing
        mstrErrorString = "Error while generating child order string." & vbCrLf & "(" & Err.Description & ")"
        DoChildOrderString = ""

    End Function

    Private Function GenerateSQLWhere() As Boolean

        ' Purpose : Generate the where clauses that cope with the joins
        '           NB Need to add the where clauses for filters/picklists etc
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo GenerateSQLWhere_ERROR

        Dim blnOK As Boolean
        Dim pintLoop As Short
        Dim pobjTableView As TablePrivilege
        Dim prstTemp As New ADODB.Recordset
        Dim pstrPickListIDs As String
        Dim strFilterIDs As String
        Dim objExpr As clsExprExpression
        Dim pstrParent1PickListIDs As String
        Dim pstrParent2PickListIDs As String

        '# remove this if test dont work RH 18/05
        pobjTableView = gcoTablePrivileges.FindTableID(mlngExportBaseTable)
        If pobjTableView.AllowSelect = False Then
            '#

            ' First put the where clauses in for the joins...only if base table is a top level table
            If UCase(Left(mstrBaseTableRealSource, 6)) <> "ASRSYS" Then

                For pintLoop = 1 To UBound(mlngTableViews, 2)
                    ' Get the table/view object from the id stored in the array
                    If mlngTableViews(1, pintLoop) = 0 Then
                        pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
                    Else
                        pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
                    End If

                    ' dont add where clause for the base/chil/p1/p2 TABLES...only add views here
                    ' JPD20030207 Fault 5033
                    'If (mlngTableViews(1, pintLoop) = 1) And _
                    '(mlngTableViews(2, pintLoop) <> mlngExportChildTable) _
                    'And (mlngTableViews(2, pintLoop) <> mlngExportParent1Table) _
                    'And (mlngTableViews(2, pintLoop) <> mlngExportParent2Table) _
                    'And (mlngTableViews(2, pintLoop) <> mlngExportBaseTable) Then
                    If (mlngTableViews(1, pintLoop) = 1) Then
                        mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " OR ", " WHERE (") & mstrBaseTableRealSource & ".ID IN (SELECT ID FROM " & pobjTableView.RealSource & ")"
                    End If

                Next pintLoop

                If Len(mstrSQLWhere) > 0 Then mstrSQLWhere = mstrSQLWhere & ")"

            End If

            '# remove this if test dont work RH 18/05
        End If
        '#

        ' Parent 1 filter and picklist
        If mlngExportParent1PickListID > 0 Then
            pstrParent1PickListIDs = ""
            prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngExportParent1PickListID, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            If prstTemp.BOF And prstTemp.EOF Then
                mstrErrorString = "The first parent table picklist contains no records."
                GenerateSQLWhere = False
                Exit Function
            End If

            Do While Not prstTemp.EOF
                pstrParent1PickListIDs = pstrParent1PickListIDs & IIf(Len(pstrParent1PickListIDs) > 0, ", ", "") & prstTemp.Fields(0).Value
                prstTemp.MoveNext()
            Loop

            prstTemp.Close()
            'UPGRADE_NOTE: Object prstTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            prstTemp = Nothing

            mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngExportParent1Table & " IN (" & pstrParent1PickListIDs & ") "
        ElseIf mlngExportParent1FilterID > 0 Then
            blnOK = True
            blnOK = datGeneral.FilteredIDs(mlngExportParent1FilterID, strFilterIDs)

            ' Generate any UDFs that are used in this filter
            If blnOK Then
                datGeneral.FilterUDFs(mlngExportParent1FilterID, mastrUDFsRequired)
            End If

            If blnOK Then
                mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngExportParent1Table & " IN (" & strFilterIDs & ") "
            Else
                mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngExportParent1FilterID) & "' filter."
                GenerateSQLWhere = False
                Exit Function
            End If
        End If

        ' Parent 2 filter and picklist
        If mlngExportParent2PickListID > 0 Then
            pstrParent2PickListIDs = ""
            prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngExportParent2PickListID, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            If prstTemp.BOF And prstTemp.EOF Then
                mstrErrorString = "The second parent table picklist contains no records."
                GenerateSQLWhere = False
                Exit Function
            End If

            Do While Not prstTemp.EOF
                pstrParent2PickListIDs = pstrParent2PickListIDs & IIf(Len(pstrParent2PickListIDs) > 0, ", ", "") & prstTemp.Fields(0).Value
                prstTemp.MoveNext()
            Loop

            prstTemp.Close()
            'UPGRADE_NOTE: Object prstTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            prstTemp = Nothing

            mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngExportParent2Table & " IN (" & pstrParent2PickListIDs & ") "
        ElseIf mlngExportParent2FilterID > 0 Then
            blnOK = True
            blnOK = datGeneral.FilteredIDs(mlngExportParent2FilterID, strFilterIDs)

            ' Generate any UDFs that are used in this filter
            If blnOK Then
                datGeneral.FilterUDFs(mlngExportParent2FilterID, mastrUDFsRequired)
            End If

            If blnOK Then
                mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID_" & mlngExportParent2Table & " IN (" & strFilterIDs & ") "
            Else
                mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngExportParent2FilterID) & "' filter."
                GenerateSQLWhere = False
                Exit Function
            End If
        End If

        ' Now if we are using a picklist, add a where clause for that
        'Get List of IDs from Picklist
        If mlngExportPickListID > 0 Then
            prstTemp = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngExportPickListID, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            If prstTemp.BOF And prstTemp.EOF Then
                mstrErrorString = "The selected picklist contains no records."
                GenerateSQLWhere = False
                Exit Function
            End If

            Do While Not prstTemp.EOF
                pstrPickListIDs = pstrPickListIDs & IIf(Len(pstrPickListIDs) > 0, ", ", "") & prstTemp.Fields(0).Value
                prstTemp.MoveNext()
            Loop

            prstTemp.Close()
            'UPGRADE_NOTE: Object prstTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            prstTemp = Nothing

            mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & pstrPickListIDs & ")"

        ElseIf mlngExportFilterID > 0 Then

            blnOK = datGeneral.FilteredIDs(mlngExportFilterID, strFilterIDs)

            ' Generate any UDFs that are used in this filter
            If blnOK Then
                datGeneral.FilterUDFs(mlngExportFilterID, mastrUDFsRequired)
            End If

            If blnOK Then
                mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrSQLFrom & ".ID IN (" & strFilterIDs & ")"
            Else
                ' Permission denied on something in the filter.
                mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngExportFilterID) & "' filter."
                GenerateSQLWhere = False
                Exit Function
            End If
        End If

        'UPGRADE_NOTE: Object prstTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        prstTemp = Nothing

        ' Audit changes only
        If Len(mstrOnlyChangesFilter) > 0 And mbAuditChangesOnly And mlngOutputFormat = modEnums.OutputFormats.fmtXML Then
            mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & "(" & mstrOnlyChangesFilter & ")"
        End If

        GenerateSQLWhere = True
        Exit Function

GenerateSQLWhere_ERROR:

        GenerateSQLWhere = False
        mstrErrorString = "Error in GenerateSQLWhere." & vbCrLf & "(" & Err.Description & ")"

    End Function

    Private Function GenerateSQLOrderBy() As Boolean

        ' Purpose : Returns order by string from the sort order array
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo GenerateSQLOrderBy_ERROR

        If UBound(mvarSortOrder, 2) > 0 Then
            ' Columns have been defined, so use these for the base table/view
            mstrSQLOrderBy = DoDefinedOrderBy()
        End If

        If Len(mstrSQLOrderBy) > 0 Then mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy


        GenerateSQLOrderBy = True
        Exit Function

GenerateSQLOrderBy_ERROR:

        GenerateSQLOrderBy = False
        mstrErrorString = "Error in GenerateSQLOrderBy." & vbCrLf & "(" & Err.Description & ")"

    End Function

    Private Function DoDefinedOrderBy() As String

        ' This function creates the base ORDER BY statement by searching
        ' through the columns defined as the reports sort order, then
        ' uses the relevant alias name

        Dim iLoop As Short
        Dim iLoop2 As Short

        For iLoop = 1 To UBound(mvarSortOrder, 2)

            For iLoop2 = 1 To UBound(mvarColDetails, 2)

                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If mvarSortOrder(1, iLoop) = mvarColDetails(3, iLoop2) Then

                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    DoDefinedOrderBy = DoDefinedOrderBy & IIf(Len(DoDefinedOrderBy) > 0, ",", "") & "[" & mvarColDetails(14, iLoop2) & "] " & mvarSortOrder(2, iLoop)

                    '"[" & mvarColDetails(5, iLoop2) & "] " & _
                    '
                    Exit For

                End If

            Next iLoop2

        Next iLoop

    End Function

    Private Function CheckRecordSet() As Boolean

        ' Purpose : To get recordset from temptable and show recordcount
        ' Input   : None
        ' Output  : True/False Success

        On Error GoTo CheckRecordSet_ERROR

        'gobjProgress.UpdateProgress gblnBatchMode

        ' If user cancels the export, abort
        If gobjProgress.Cancelled Then
            mblnUserCancelled = True
            CheckRecordSet = False
            Exit Function
        End If

        mrstExportOutput = mclsData.OpenRecordset("SELECT * FROM " & mstrTempTableName, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If mrstExportOutput.BOF And mrstExportOutput.EOF Then

            ' If we want to force the header trick export into thinking we have records
            If mbForceHeader Then
                CheckRecordSet = True
            Else
                CheckRecordSet = False
                mstrErrorString = "Export generated no records !"
            End If

            mblnNoRecords = True
            Exit Function
        End If

        mrstExportOutput.MoveLast()
        mrstExportOutput.MoveFirst()

        mlngExportRecordCount = mrstExportOutput.RecordCount

        If gblnBatchMode = False Then
            gobjProgress.Bar1MaxValue = mlngExportRecordCount
        Else
            gobjProgress.Bar2MaxValue = mlngExportRecordCount
        End If

        CheckRecordSet = True
        Exit Function

CheckRecordSet_ERROR:

        mstrErrorString = "Error while checking returned recordset." & vbCrLf & "(" & Err.Description & ")"
        CheckRecordSet = False

    End Function

    Private Function ClearUp() As Boolean

        ' Purpose : To clear all variables/recordsets/references and drops temptable
        ' Input   : None
        ' Output  : True/False success

        ' Definition variables

        On Error GoTo ClearUp_ERROR

        mlngExportID = 0
        mstrExportName = vbNullString
        mstrExportDescription = vbNullString
        mlngExportBaseTable = 0
        mstrExportBaseTableName = vbNullString
        mlngExportAllRecords = 1
        mlngExportPickListID = 0
        mlngExportFilterID = 0
        mlngExportParent1Table = 0
        mstrExportParent1TableName = vbNullString
        mlngExportParent1FilterID = 0
        mlngExportParent2Table = 0
        mstrExportParent2TableName = vbNullString
        mlngExportParent2FilterID = 0
        mlngExportChildTable = 0
        mstrExportChildTableName = vbNullString
        mlngExportChildFilterID = 0
        mlngExportChildMaxRecords = 0
        'mstrExportOutputType = vbNullString
        'mstrExportOutputName = vbNullString
        mblnExportQuotes = False
        'mblnExportHeader = False
        mintExportHeader = 0
        mstrExportHeaderText = vbNullString
        mintExportFooter = 0
        mstrExportFooterText = vbNullString

        mstrExportDateFormat = vbNullString
        'mstrExportDateSeparator = vbNullString
        'mstrExportDateYearDigits = vbNullString
        mlngExportRecordCount = 0
        'mstrErrorString = vbNullString

        mlngExportParent1AllRecords = 1
        mlngExportParent1PickListID = 0
        mlngExportParent2AllRecords = 1
        mlngExportParent2PickListID = 0

        'mstrErrorString = vbNullString
        mblnNoRecords = False
        mlSuccessfulRecords = 0
        mlngExportRecordCount = 0

        ' Recordsets

        'UPGRADE_NOTE: Object mrstExportOutput may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        mrstExportOutput = Nothing

        'TM20020531 Fault 3756
        ' Delete the temptable if exists, and then clear the variable
        '  If Len(mstrTempTableName) > 0 Then
        '
        '    mclsData.ExecuteSql ("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & _
        ''                         "DROP TABLE " & mstrTempTableName)
        '  End If
        datGeneral.DropUniqueSQLObject(mstrTempTableName, 3)
        mstrTempTableName = vbNullString

        ' SQL strings

        mstrSQLSelect = vbNullString
        mstrSQLFrom = vbNullString
        mstrSQLWhere = vbNullString
        mstrSQLJoin = vbNullString
        mstrSQLOrderBy = vbNullString
        mstrSQL = vbNullString

        ' Class references

        'Set mclsData = Nothing
        'Set mclsGeneral = Nothing

        ' Arrays

        ReDim mvarColDetails(UBound(mvarColDetails, 1), 0)
        ReDim mvarSortOrder(2, 0)

        ' Column Privilege arrays / collections / variables

        mstrBaseTableRealSource = vbNullString
        mstrRealSource = vbNullString
        'Set mobjTableView = Nothing
        'Set mobjColumnPrivileges = Nothing
        ReDim mlngTableViews(2, 0)
        ReDim mstrViews(0)

        gobjProgress.ResetBar2()
        mblnUserCancelled = False

        ClearUp = True

        Exit Function

ClearUp_ERROR:

        mstrErrorString = "Error whilst clearing data." & vbCrLf & "(" & Err.Description & ")"
        ClearUp = False

    End Function

    Private Function ExportData() As Boolean

        On Error GoTo ExportData_ERROR

        'gobjProgress.UpdateProgress gblnBatchMode

        mlSuccessfulRecords = 0

        ' If user cancels the export, abort
        If gobjProgress.Cancelled Then
            mblnUserCancelled = True
            ExportData = False
            Exit Function
        End If

        Select Case mlngOutputFormat
            Case modEnums.OutputFormats.fmtXML
                ExportData = ExportData_XML()

            Case modEnums.OutputFormats.fmtCMGFile
                ExportData = ExportData_CMGfile()

            Case Else
                If ReadDataIntoArray() Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object SendArrayToOutputOptions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ExportData = SendArrayToOutputOptions()
                End If

        End Select

        Exit Function

ExportData_ERROR:

        mstrErrorString = "Error whilst exporting data." & vbCrLf & "(" & Err.Description & ")"
        ExportData = False

    End Function

    'Private Function ExportData_Delimited() As Boolean
    '
    '  ' The export is to an delimited file.
    '
    '  Dim pintFileNo As Integer
    '  Dim pstrExportString As String
    '  'Dim pstrDateString As String
    '  Dim prstReadyToExport As Recordset
    '  Dim pintLoop As Integer
    '  Dim lngRecordNumber As Long
    '  Dim bAppended As Boolean
    '  Dim bForceHeader As Boolean
    '  Dim bNoRecords As Boolean
    '
    '  Dim tmpDec As Long
    '  Dim tmpLen As Long
    '
    '  Dim objExpr As clsExprExpression
    '
    '  On Error GoTo ExportData_Delimited_ERROR
    '
    '  ' If filename specified already exists then delete it first.
    '  If Len(Dir(mstrOutputFilename)) > 0 Then
    '    If mlngOutputSaveExisting = 4 Then
    '      bAppended = True
    '    Else
    '      bAppended = False
    '      Kill mstrOutputFilename
    '    End If
    '  Else
    '    bAppended = False
    '  End If
    '
    '  ' Open file for output.
    '  pintFileNo = FreeFile
    '
    '  If bAppended Then
    '    Open mstrOutputFilename For Append As pintFileNo
    '  Else
    '    Open mstrOutputFilename For Output As pintFileNo
    '  End If
    '
    '  ' Open the export table as a recordset.
    '  Set prstReadyToExport = mclsGeneral.GetRecords("SELECT * FROM [" & mstrTempTableName & "]")
    '
    '  If (prstReadyToExport.BOF And prstReadyToExport.EOF) Then
    '    mstrErrorString = "No records to export."
    '    mclsData.ExecuteSql ("DROP TABLE [" & mstrTempTableName & "]")
    '    ExportData_Delimited = False
    '    bForceHeader = mbForceHeader
    '    bNoRecords = True
    '  Else
    '    bForceHeader = False
    '    bNoRecords = False
    '  End If
    '
    '  ' Omit the header if we are appending to file (if specified)
    '  If Not (bAppended And mbOmitHeader) Or bForceHeader Then
    '
    '    Select Case mintExportHeader
    '    Case 1  'Column Names
    '      'MH20030120
    '      'pstrExportString = vbNullString
    '      'For pintLoop = 0 To prstReadyToExport.Fields.Count - 1
    '      '  pstrExportString = pstrExportString & IIf(Len(pstrExportString) > 0, ",", "") & prstReadyToExport.Fields(pintLoop).Name
    '      'Next pintLoop
    '      pstrExportString = vbNullString
    '      For pintLoop = 0 To UBound(mvarColDetails, 2)
    '        pstrExportString = pstrExportString & _
    ''          IIf(pstrExportString <> vbNullString, ",", vbNullString) & _
    ''          mvarColDetails(13, pintLoop)
    '      Next
    '      Print #pintFileNo, pstrExportString
    '    Case 2  'Custom Heading
    '      Print #pintFileNo, mstrExportHeaderText
    '    End Select
    '
    '  End If
    '
    '  ' If we only need to output the header bomb out here
    '  If bNoRecords Then
    '    Close #pintFileNo
    '    Exit Function
    '  End If
    '
    '  ' Loop through the export table and print stuff to file.
    '  lngRecordNumber = 0
    '  Do While Not prstReadyToExport.EOF
    '
    '    ' If user cancels the export, abort
    '    If gobjProgress.Cancelled Then
    '      mblnUserCancelled = True
    '      Close #pintFileNo
    '      ExportData_Delimited = False
    '      Exit Function
    '    End If
    '
    '    pstrExportString = vbNullString
    '
    '    ' Loop through the fields in each record, adding them and the delimiter to the export string.
    '    lngRecordNumber = lngRecordNumber + 1
    '    For pintLoop = 0 To prstReadyToExport.Fields.Count - 1
    '
    '
    '      'MH20000705 Fault 541
    '      'When exporting a sort code it was being converted to a date !!??!??!
    '      'so check if its a column, if yes then check if it is a date column
    '
    '      '' IsDate isnt perfect, so check for a ':' char too...if its there, it is a date
    '      'If IsDate(prstReadyToExport.Fields(pintLoop)) And InStr(prstReadyToExport.Fields(pintLoop), ":") = 0 Then
    '
    '      If mvarColDetails(0, pintLoop + 1) = "N" Then
    '        'Record Number
    '        pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & CStr(lngRecordNumber) & _
    ''        IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
    '
    '      ElseIf mvarColDetails(7, pintLoop + 1) = True Then
    '
    '
    '        'pstrExportString = pstrExportString & Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat) & _
    ''                           IIf(mstrExportDelimiter <> ",", vbTab, ",")
    '
    '
    '              If Len(Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat)) > mvarColDetails(6, pintLoop + 1) Then
    ''                pstrExportString = pstrExportString & Left(Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat), mvarColDetails(6, pintLoop + 1)) & _
    ''                IIf(mstrExportDelimiter <> ",", vbTab, ",")
    '                pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat), mvarColDetails(6, pintLoop + 1)) & _
    ''                IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
    '              Else
    ''                pstrExportString = pstrExportString & Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat) & _
    ''                IIf(mstrExportDelimiter <> ",", vbTab, ",")
    '                pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Format(prstReadyToExport.Fields(pintLoop), mstrExportDateFormat) & _
    ''                IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
    '              End If
    '
    '
    '      ElseIf mvarColDetails(12, pintLoop + 1) Then
    '
    '        '#RH 04/11
    '        ' Request from m edwynn - trim the data when exporting to a csv
    '
    ''        pstrExportString = pstrExportString & Trim(prstReadyToExport.Fields(pintLoop)) & _
    ''                           IIf(mstrExportDelimiter <> ",", vbTab, ",")
    '
    '                'TM20011010 Fault 2197
    '                'When creating the string, use the FormatNumeric() function so the
    '                'Size and Decimals of the numeric can be formatted as required.
    '                tmpLen = mvarColDetails(6, pintLoop + 1)
    '                tmpDec = mvarColDetails(11, pintLoop + 1)
    '
    '                'TM20020507 Fault 3840 - Stupid NULLs mucked it up.
    '                pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & FormatNumeric(IIf(IsNull(prstReadyToExport.Fields(pintLoop)), 0, prstReadyToExport.Fields(pintLoop)), tmpLen, tmpDec) & _
    ''                  IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
    '
    ''                If Len(prstReadyToExport.Fields(pintLoop)) > mvarColDetails(6, pintLoop + 1) Then
    '''                  pstrExportString = pstrExportString & Left(prstReadyToExport.Fields(pintLoop), mvarColDetails(6, pintLoop + 1)) & _
    '''                  IIf(mstrExportDelimiter <> ",", vbTab, ",")
    ''                  pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & FormatNumeric(prstReadyToExport.Fields(pintLoop), tmpLen, tmpDec) & _
    '''                  IIf(mblnExportQuotes, Chr(34), "") & IIf(mstrExportDelimiter <> ",", vbTab, ",")
    ''                Else
    '''                  pstrExportString = pstrExportString & prstReadyToExport.Fields(pintLoop) & _
    '''                  IIf(mstrExportDelimiter <> ",", vbTab, ",")
    ''                  pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & FormatNumeric(prstReadyToExport.Fields(pintLoop), tmpLen, tmpDec) & _
    '''                  IIf(mblnExportQuotes, Chr(34), "") & IIf(mstrExportDelimiter <> ",", vbTab, ",")
    ''                End If
    '
    '
    '      Else
    '
    '         'RH - BUG 614
    '
    '        If Len(prstReadyToExport.Fields(pintLoop)) > mvarColDetails(6, pintLoop + 1) Then
    '
    '          ' Need to trim the data, its longer than the length required to export
    '
    '          ' Trim is here
    ''          pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(Trim(prstReadyToExport.Fields(pintLoop)), mvarColDetails(6, pintLoop + 1)) & _
    ''           IIf(mblnExportQuotes, Chr(34), "") & IIf(mstrExportDelimiter <> ",", vbTab, ",")
    '
    '          ' Take away the trim - ruins working pattern field !
    '          pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(prstReadyToExport.Fields(pintLoop), mvarColDetails(6, pintLoop + 1)) & _
    ''           IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
    '
    '        Else
    '
    '          ' Data is fine as it is (dont pad with spaces for delimited files!)
    '
    '          ' Trim is here
    '
    ''          pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(Trim(prstReadyToExport.Fields(pintLoop)), mvarColDetails(6, pintLoop + 1)) & _
    ''           IIf(mblnExportQuotes, Chr(34), "") & IIf(mstrExportDelimiter <> ",", vbTab, ",")
    '
    '          ' Take away the trim - ruins working pattern field !!!
    '          pstrExportString = pstrExportString & IIf(mblnExportQuotes, Chr(34), "") & Left(prstReadyToExport.Fields(pintLoop), mvarColDetails(6, pintLoop + 1)) & _
    ''           IIf(mblnExportQuotes, Chr(34), "") & mstrExportDelimiter
    '
    '        End If
    '
    '      End If
    '
    '    Next pintLoop
    '
    '    ' If the delimiter is "," then remove the last character.
    '    If Not mstrExportDelimiter = vbTab Then
    '      pstrExportString = Left(pstrExportString, Len(pstrExportString) - 1)
    '    End If
    '
    '    ' Print the stuff to the file.
    '    Print #pintFileNo, pstrExportString
    '
    '    'JDM - 13/12/01 - Fault 3280 - Log successful records
    '    If mbLoggingExportSuccess Then
    '      gobjEventLog.AddDetailEntry pstrExportString & " Exported successfully"
    '    End If
    '
    '    mlSuccessfulRecords = mlSuccessfulRecords + 1
    '
    '    prstReadyToExport.MoveNext
    '
    '  Loop
    '
    '
    '  Select Case mintExportFooter
    '  Case 1  'Column Names
    '    pstrExportString = vbNullString
    '    For pintLoop = 0 To prstReadyToExport.Fields.Count - 1
    '      pstrExportString = pstrExportString & IIf(Len(pstrExportString) > 0, ",", "") & prstReadyToExport.Fields(pintLoop).Name
    '    Next pintLoop
    '    Print #pintFileNo, pstrExportString
    '  Case 2  'Custom Heading
    '    Print #pintFileNo, mstrExportFooterText
    '  End Select
    '
    '
    '  ' Close the final output file
    '  Close #pintFileNo
    '
    '  'Drop the export table as we dont need it now.
    '  mclsData.ExecuteSql ("DROP TABLE [" & mstrTempTableName & "]")
    '
    '  Set prstReadyToExport = Nothing
    '  ExportData_Delimited = True
    '  Exit Function
    '
    'ExportData_Delimited_ERROR:
    '
    '  Select Case Err.Number
    '    Case 70 ' file creation error / sharing violation
    '      mstrErrorString = "Error creating file '" & mstrOutputFilename & "'. File may already be in use." & vbCrLf & "(" & Err.Description & ")"
    '    Case 76 ' path not found error
    '      mlngExportRecordCount = 0
    '      mstrErrorString = "Error whilst exporting to delimited file." & vbCrLf & "(" & Err.Description & ")"
    '    Case Else
    '    mstrErrorString = "Error whilst exporting to delimited file." & vbCrLf & "(" & Err.Description & ")"
    '    Close #pintFileNo
    '  End Select
    '
    '  ExportData_Delimited = False
    '
    'End Function

    Private Function FormatNumeric(ByRef dValue As Double, ByRef iLen As Integer, ByRef iDec As Integer) As Object

        Const sDecimalChar As String = "."

        Dim iCount As Short
        Dim iDecimalCount As Short
        Dim iLengthCount As Short
        Dim sCurrDig As String
        Dim bHasDec As Boolean
        Dim sDecimalFormat As String
        Dim i As Short

        iCount = 0
        iDecimalCount = 0
        iLengthCount = 0
        sCurrDig = vbNullString
        bHasDec = False
        sDecimalFormat = vbNullString

        If iLen > 0 Then
            Do While iCount <= iLen
                iCount = iCount + 1
                sCurrDig = Mid(CStr(dValue), iCount, 1)

                If sCurrDig = sDecimalChar Then
                    bHasDec = True
                Else
                    If bHasDec And iDecimalCount < iDec Then
                        iDecimalCount = iDecimalCount + 1
                        iLengthCount = iLengthCount + 1
                    ElseIf Not bHasDec Then
                        iLengthCount = iLengthCount + 1
                    End If
                End If
            Loop

            'TM20020417 Fault 3772
            If iDec > 0 Then
                For i = 1 To iDec Step 1
                    sDecimalFormat = sDecimalFormat & "0"
                Next i
                'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
                If LenB(sDecimalFormat) <> 0 Then
                    sDecimalFormat = "###0." & sDecimalFormat
                End If
                FormatNumeric = VB6.Format(System.Math.Round(dValue, iDecimalCount), sDecimalFormat)

            Else
                FormatNumeric = System.Math.Round(dValue, iDecimalCount)
            End If

        Else
            If iDecimalCount > 0 Then
                FormatNumeric = System.Math.Round(dValue, iDecimalCount)
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object FormatNumeric. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                FormatNumeric = CStr(dValue)
            End If
        End If

    End Function

    Public Function PrettyPrintXml(ByRef xmlDoc As Object) As Object
        Dim rdr, wrt As Object
        rdr = CreateObject("Msxml2.SAXXMLReader.3.0")
        wrt = CreateObject("Msxml2.MXXMLWriter.3.0")
        'UPGRADE_WARNING: Couldn't resolve default property of object wrt.Encoding. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        wrt.Encoding = "UTF-8"
        'UPGRADE_WARNING: Couldn't resolve default property of object wrt.standalone. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        wrt.standalone = False
        'UPGRADE_WARNING: Couldn't resolve default property of object wrt.byteOrderMark. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        wrt.byteOrderMark = True
        'UPGRADE_WARNING: Couldn't resolve default property of object wrt.omitXMLDeclaration. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        wrt.omitXMLDeclaration = False
        'UPGRADE_WARNING: Couldn't resolve default property of object wrt.Indent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        wrt.Indent = True

        'UPGRADE_WARNING: Couldn't resolve default property of object rdr.contentHandler. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rdr.contentHandler = wrt
        'UPGRADE_WARNING: Couldn't resolve default property of object rdr.dtdHandler. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rdr.dtdHandler = wrt
        'UPGRADE_WARNING: Couldn't resolve default property of object rdr.errorHandler. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rdr.errorHandler = wrt
        'UPGRADE_WARNING: Couldn't resolve default property of object rdr.PutProperty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rdr.PutProperty("http://xml.org/sax/properties/lexical-handler", wrt)
        'UPGRADE_WARNING: Couldn't resolve default property of object rdr.PutProperty. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rdr.PutProperty("http://xml.org/sax/properties/declaration-handler", wrt)
        'UPGRADE_WARNING: Couldn't resolve default property of object rdr.Parse. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        rdr.Parse(xmlDoc)

        Dim sRemoveHeader As String
        sRemoveHeader = "standalone=""no"""
        'UPGRADE_WARNING: Couldn't resolve default property of object wrt.Output. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrettyPrintXml = Replace(wrt.Output, sRemoveHeader, "")
        PrettyPrintXml = Replace(PrettyPrintXml, "encoding=""UTF-16""", "encoding=""UTF-8""")

    End Function

    Private Function ExportData_XML() As Boolean

        Dim pintFileNo As Short
        Dim bAppended As Boolean
        Dim sExportText As String
        Dim sTransFormText As String
        Dim sSQL As String
        Dim sSQLGetData As String
        Dim prstReadyToExport As ADODB.Recordset
        Dim rstConverted As ADODB.Recordset
        Dim iFileCount As Short
        Dim strOutputFilename As String
        Dim sLimitedColumns As String
        Dim iCount As Short

        Dim sXSDName As String
        Dim sXSDTransformFile As String
        Dim sWhereClause As String

        strOutputFilename = mstrOutputFileName
        sXSDName = mstrXSDFilename
        sXSDTransformFile = mstrTransformFile

        If Not mbPreserveTransformPath And Len(mstrTransformFile) > 0 Then
            sXSDTransformFile = GetFileNameOnly(mstrTransformFile)
        End If

        If Not mbPreserveXSDPath And Len(mstrXSDFilename) > 0 Then
            sXSDName = GetFileNameOnly(mstrXSDFilename)
        End If

        ' Limit the column selection to actual columns the user has selected (XML has nto include the ID for the splitter to work)
        sLimitedColumns = ""
        For iCount = 1 To UBound(mvarColDetails, 2) - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(14, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sLimitedColumns = sLimitedColumns & IIf(Len(sLimitedColumns) > 0, ", ", "") & "[" & mvarColDetails(14, iCount) & "]"
        Next

        If mbSplitXMLIntoFiles Then
            sWhereClause = " WHERE basetable._XMLSplitID = _XMLSplitID "
        End If

        sSQLGetData = "SELECT CAST((SELECT " & sLimitedColumns & " FROM [" & mstrTempTableName & "]" & sWhereClause & " FOR XML PATH('" & mstrXMLDataNodeName & "'), ROOT('" & mstrExportHeaderText & "')" & ", ELEMENTS XSINIL) AS VARCHAR(MAX)) AS XmlData"

        If mbSplitXMLIntoFiles Then
            sSQLGetData = "SELECT XMLData =(" & sSQLGetData & ") FROM [" & mstrTempTableName & "] basetable"
        End If

        If Len(sXSDName) > 0 Then
            sSQLGetData = "WITH XMLNAMESPACES ('" & sXSDName & "' as noNamespaceSchemaLocation)" & sSQLGetData
        End If

        prstReadyToExport = mclsData.OpenRecordset(sSQLGetData, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If (prstReadyToExport.BOF And prstReadyToExport.EOF) Then
            mstrErrorString = "No records to export."
            mclsData.ExecuteSql(("IF EXISTS(SELECT Name FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & "DROP TABLE " & mstrTempTableName))
            ExportData_XML = False
            Exit Function
        Else
            prstReadyToExport.MoveFirst()

            Do While Not prstReadyToExport.EOF

                iFileCount = iFileCount + 1

                If Len(sXSDTransformFile) > 0 Then
                    sExportText = prstReadyToExport.Fields(0).value
                    sTransFormText = GetFileText(sXSDTransformFile)
                    sSQL = "SELECT convert(nvarchar(MAX), dbo.udfASRNetApplyXsltTransform(convert(xml,'" & Replace(sExportText, "'", "''") & "'), convert(xml,'" & Replace(sTransFormText, "'", "''") & "')))"
                    rstConverted = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                    'UPGRADE_WARNING: Couldn't resolve default property of object PrettyPrintXml(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    sExportText = PrettyPrintXml((rstConverted.Fields(0).value))
                Else
                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                    If Not IsDBNull(prstReadyToExport.Fields(0).value) Then
                        sExportText = prstReadyToExport.Fields(0).value
                        'UPGRADE_WARNING: Couldn't resolve default property of object PrettyPrintXml(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        sExportText = PrettyPrintXml(sExportText)

                        If Len(sXSDName) > 0 Then
                            sExportText = Replace(sExportText, "xmlns:noNamespaceSchemaLocation", "xsi:noNamespaceSchemaLocation")
                        End If

                    End If
                End If

                If mbSplitXMLIntoFiles Then
                    strOutputFilename = InsertNumberIntoFilename(mstrOutputFileName, iFileCount)
                End If

                ' Process the file name
                strOutputFilename = ReplaceFormatExpressions(strOutputFilename, 1, mlngExportRecordCount)

                Select Case mlngOutputSaveExisting
                    Case 0 ' Overwrite
                        bAppended = False
                    ' Kill mstrOutputFileName

                    Case 1 ' Do not overwrite (fail)
                        mstrErrorString = "File already exists."
                        ExportData_XML = False
                        Exit Function

                    Case 2 ' Add sequential number to filename
                        mstrOutputFileName = GetSequentialNumberedFile(strOutputFilename)
                        bAppended = False

                    Case 3 ' Append to Existing File
                        bAppended = True

                End Select

                ' Open file for output.
                pintFileNo = FreeFile()

                If bAppended Then
                    FileOpen(pintFileNo, strOutputFilename, OpenMode.Append)
                Else
                    FileOpen(pintFileNo, strOutputFilename, OpenMode.Output)
                End If

                PrintLine(pintFileNo, sExportText)

                FileClose(pintFileNo)

                prstReadyToExport.MoveNext()
            Loop

        End If

        ExportData_XML = True

    End Function


    Private Function ExportData_CMGfile() As Boolean

        'JDM - 21/08/01 - Fault 2706 - Not to overwrite file if no records found

        ' The export is to a CMG file.

        Dim iLoop As Short
        Dim pintFileNo As Short
        Dim pstrExportString As DataMgr.clsStringBuilder
        Dim strExportColumn As String
        Dim prstReadyToExport As ADODB.Recordset
        Dim pintLoop As Short
        Dim bOutputColumn As Boolean
        Dim dLastChangeDate As Date
        Dim strRecordIdentifier As String
        Dim bCMGExportFileCode As Boolean
        Dim bCMGExportFieldCode As Boolean
        Dim bCMGExportLastChangeDate As Boolean
        Dim iCMGExportFileCodeSize As Short
        Dim iCMGEXportRecordIDSize As Short
        Dim iCMGExportFieldCodeSize As Short
        Dim iCMGExportOutputColumnSize As Short
        Dim iCMGExportLastChangeDateSize As Short
        'NPG20090313 Fault 13595
        Dim iCMGExportFileCodeOrderID As Short
        Dim iCMGEXportRecordIDOrderID As Short
        Dim iCMGExportFieldCodeOrderID As Short
        Dim iCMGExportOutputColumnOrderID As Short
        Dim iCMGExportLastChangeDateOrderID As Short

        Dim bExportFileOpened As Boolean
        Dim bUseCSV As Boolean
        Dim bCMGIgnoreBlanks As Boolean
        Dim bCMGReverseOutput As Boolean
        Dim iCount As Short
        Dim bAppended As Boolean

        Dim lngRecordCount As Integer
        Dim astrColumnIDs() As String
        Dim astrRecordIDs() As String
        Dim astrBulkRecordIDs() As String
        'NPG20071218 Fault 12867
        Dim astrConvertCase() As String
        'NPG20080617 Suggestion S000816
        Dim astrSuppressNulls() As String
        Dim strRecordIDs As String
        Dim strColumnIDs As String
        'NPG20071218 Fault 12867
        Dim strConvertCase As String
        'NPG20080617 Suggestion S000816
        Dim strSuppressNulls As String
        Dim sSQL As String
        Dim rsAffectedRecords As ADODB.Recordset
        Dim varNewValue As String
        Dim astrColumnCodes() As String
        Dim aiDataTypes() As Short
        Dim pbNullSuppressed As Boolean
        'Dim lngLastColumnID As Long
        Dim strCMGCode As String
        Dim iDataType As Short
        Dim fNewLine As Boolean

        Dim iRow As Short
        Dim fNewRow As Boolean

        On Error GoTo ExportData_CMGfile_ERROR

        pstrExportString = New DataMgr.clsStringBuilder
        rsAffectedRecords = New ADODB.Recordset

        ' Which fields to export and their size
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bUseCSV = GetSystemSetting("CMGExport", "UseCSV", False)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bCMGIgnoreBlanks = GetSystemSetting("CMGExport", "IgnoreBlanks", False)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bCMGReverseOutput = GetSystemSetting("CMGExport", "ReverseOutput", False)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bCMGExportFileCode = GetSystemSetting("CMGExport", "FileCode", True)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bCMGExportFieldCode = GetSystemSetting("CMGExport", "FieldCode", True)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        bCMGExportLastChangeDate = GetSystemSetting("CMGExport", "LastChange", True)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGExportFileCodeSize = GetSystemSetting("CMGExport", "FileCodeSize", 6)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGEXportRecordIDSize = GetSystemSetting("CMGExport", "RecordIdentifierSize", 11)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGExportFieldCodeSize = GetSystemSetting("CMGExport", "FieldCodeSize", 10)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGExportOutputColumnSize = GetSystemSetting("CMGExport", "OutputColumnSize", 53)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGExportLastChangeDateSize = GetSystemSetting("CMGExport", "LastChangeSize", 8)
        'NPG20090313 Fault 13595
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGExportFileCodeOrderID = GetSystemSetting("CMGExport", "FileCodeOrderID", 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGEXportRecordIDOrderID = GetSystemSetting("CMGExport", "RecordIdentifierOrderID", 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGExportFieldCodeOrderID = GetSystemSetting("CMGExport", "FieldCodeOrderID", 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGExportOutputColumnOrderID = GetSystemSetting("CMGExport", "OutputColumnOrderID", 0)
        'UPGRADE_WARNING: Couldn't resolve default property of object GetSystemSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iCMGExportLastChangeDateOrderID = GetSystemSetting("CMGExport", "LastChangeOrderID", 0)

        bExportFileOpened = False

        ' Constantly exported strings
        If bUseCSV Then
            'NPG20090403 Fault 13636
            'mstrExportFileCode = mstrExportFileCode
        Else
            mstrExportFileCode = SetStringLength(mstrExportFileCode, iCMGExportFileCodeSize)
        End If

        ' Build a list of all the columnids we'll be searching on
        ' Use 2 arrays, because we'll be using the join statement later on in the code
        ReDim astrColumnIDs(UBound(mvarColDetails, 2))
        ReDim astrColumnCodes(UBound(mvarColDetails, 2))
        ReDim aiDataTypes(UBound(mvarColDetails, 2))
        'NPG20071218 Fault 12867
        ReDim astrConvertCase(UBound(mvarColDetails, 2))
        'NPG20080617 Suggestion S000816
        ReDim astrSuppressNulls(UBound(mvarColDetails, 2))

        iCount = 0
        For pintLoop = 0 To UBound(mvarColDetails, 2)
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(9, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If mvarColDetails(9, pintLoop) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(3, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                astrColumnIDs(iCount) = mvarColDetails(3, pintLoop) ' Column ID
                aiDataTypes(iCount) = datGeneral.GetColumnDataType(CInt(astrColumnIDs(iCount)))
                'NPG20071218 Fault 12867
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(16, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                astrConvertCase(iCount) = mvarColDetails(16, pintLoop)
                'NPG20080617 Suggestion S000816
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(17, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                astrSuppressNulls(iCount) = mvarColDetails(17, pintLoop)

                If bUseCSV Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(8, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    astrColumnCodes(iCount) = mvarColDetails(8, pintLoop)
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    astrColumnCodes(iCount) = SetStringLength(mvarColDetails(8, pintLoop), iCMGExportFieldCodeSize)
                End If

                iCount = iCount + 1
            End If
        Next pintLoop


        If iCount = 0 Then
            mstrErrorString = "The export definition contains no columns which are marked for audit."
            ExportData_CMGfile = False
            Exit Function
        End If

        ReDim Preserve astrColumnCodes(iCount - 1)
        ReDim Preserve astrColumnIDs(iCount - 1)
        ReDim Preserve aiDataTypes(iCount - 1)
        'NPG20071218 Fault 12867
        ReDim Preserve astrConvertCase(iCount - 1)
        'NPG20080617 Suggestion S000816
        ReDim Preserve astrSuppressNulls(iCount - 1)

        ' Open the export table as a recordset.
        prstReadyToExport = mclsData.OpenRecordset("SELECT ID,Identifier FROM [" & mstrTempTableName & "]", ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If (prstReadyToExport.BOF And prstReadyToExport.EOF) Then
            mstrErrorString = "No records to export."
            mclsData.ExecuteSql(("IF EXISTS(SELECT Name FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & "DROP TABLE " & mstrTempTableName))
            ExportData_CMGfile = False
            Exit Function
        End If

        ReDim astrRecordIDs(0)
        ReDim astrBulkRecordIDs(0)
        lngRecordCount = 0
        mlngExportRecordCount = 0
        strColumnIDs = Join(astrColumnIDs, ",")

        ' Loop through the export table and print stuff to file.
        Do While Not prstReadyToExport.EOF

            If lngRecordCount > UBound(astrRecordIDs) Then ReDim Preserve astrRecordIDs(lngRecordCount + 100)
            astrRecordIDs(lngRecordCount) = prstReadyToExport.Fields(0).value

            If ((lngRecordCount + 1) Mod 1000) = 0 Then
                ReDim Preserve astrBulkRecordIDs(UBound(astrBulkRecordIDs) + 1)
            Else
                astrBulkRecordIDs(UBound(astrBulkRecordIDs)) = astrBulkRecordIDs(UBound(astrBulkRecordIDs)) & IIf(lngRecordCount = 0, "", ",")
            End If

            astrBulkRecordIDs(UBound(astrBulkRecordIDs)) = astrBulkRecordIDs(UBound(astrBulkRecordIDs)) & CStr(prstReadyToExport.Fields(0).value)

            ' If user cancels the export, abort
            If gobjProgress.Cancelled Then
                mblnUserCancelled = True
                FileClose(pintFileNo)
                ExportData_CMGfile = False
                Exit Function
            Else
                If lngRecordCount Mod 50 = 0 Then
                    gobjProgress.Bar1Caption = "Calculating Audit Records ( " & Trim(Str(lngRecordCount)) & " of " & Trim(Str(prstReadyToExport.RecordCount)) & " )"
                End If
            End If

            ' Get audit transaction for this record
            sSQL = "SELECT ISNULL(NewValue, ''), DateTimeStamp, ColumnID " & "FROM ASRSysAuditTrail WHERE ID IN(SELECT MAX(ID) FROM ASRSysAuditTrail " & " WHERE ColumnID IN (" & strColumnIDs & ") And RecordID = " & astrRecordIDs(lngRecordCount) & " AND CMGCommitDate IS null AND ISNULL(OldValue,'') <> ISNULL(NewValue,'') AND NOT (OldValue = '* New Record *' AND NewValue = '')" & " GROUP BY ColumnID)"
            rsAffectedRecords.Open(sSQL, gADOCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            ' Loop through each of the returned audit records
            pstrExportString.TheString = vbNullString

            If Not rsAffectedRecords.EOF Then
                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                strRecordIdentifier = IIf(Not IsDBNull(prstReadyToExport.Fields(1).value), prstReadyToExport.Fields(1).value, vbNullString)

                If bUseCSV Then
                    ' strRecordIdentifier = strRecordIdentifier & ","
                Else
                    strRecordIdentifier = SetStringLength(strRecordIdentifier, iCMGEXportRecordIDSize)
                End If
            End If

            Do While Not rsAffectedRecords.EOF
                pbNullSuppressed = False
                'NPG20080617 Suggestion S000816
                ' This for/next block has been moved above the [File Export Code] and [Record Identifier] sections
                ' so that we can exclude any null values completely - otherwise the identifier goes into the export file.

                For iCount = 0 To UBound(astrColumnIDs, 1)
                    If astrColumnIDs(iCount) = rsAffectedRecords.Fields(CMGFields.ColumnID).value Then
                        strCMGCode = astrColumnCodes(iCount)
                        iDataType = aiDataTypes(iCount)
                        'NPG20071218 Fault 12867
                        strConvertCase = astrConvertCase(iCount)
                        'NPG20080617 Suggestion S000816
                        strSuppressNulls = astrSuppressNulls(iCount)
                        Exit For
                    End If
                Next iCount

                'NPG20080617 Suggestion S000816
                If strSuppressNulls And (rsAffectedRecords.Fields(CMGFields.NewValue).value = vbNullString Or (Left(rsAffectedRecords.Fields(CMGFields.NewValue).value, 2) = "0." And Val(rsAffectedRecords.Fields(CMGFields.NewValue).value) = CDbl("0")) Or rsAffectedRecords.Fields(CMGFields.NewValue).value = "0" Or rsAffectedRecords.Fields(CMGFields.NewValue).value = "00:00" Or rsAffectedRecords.Fields(CMGFields.NewValue).value = "False") Then
                    pbNullSuppressed = True
                End If

                'NPG20080731 Fault 13305
                If fNewLine And Not pbNullSuppressed Then pstrExportString.Append(vbNewLine)

                fNewLine = False

                'NPG20090313 Fault 13595
                If Not pbNullSuppressed Then

                    'NPG20090403 Fault 13636
                    fNewRow = True

                    For iRow = 0 To 4
                        If iCMGExportFileCodeOrderID + iCMGEXportRecordIDOrderID + iCMGExportFieldCodeOrderID + iCMGExportOutputColumnOrderID + iCMGExportLastChangeDateOrderID = 0 And iRow > 0 Then Exit For

                        ' File Export Code (padded to specified length)
                        'NPG20090313 Fault 13595
                        ' If bCMGExportFileCode Then pstrExportString.Append mstrExportFileCode
                        If iCMGExportFileCodeOrderID = iRow And bCMGExportFileCode Then
                            ' Insert the delimeter if required.
                            If bUseCSV And Not fNewRow Then pstrExportString.Append(",") ' may need to set fnewrow to false here...
                            pstrExportString.Append(mstrExportFileCode)
                            fNewRow = False
                        End If

                        ' Record Identifier (usually staff number)
                        'NPG20090313 Fault 13595
                        ' pstrExportString.Append strRecordIdentifier
                        If iCMGEXportRecordIDOrderID = iRow Then
                            ' Insert the delimeter if required.
                            If bUseCSV And Not fNewRow Then pstrExportString.Append(",") ' may need to set fnewrow to false here...
                            pstrExportString.Append(strRecordIdentifier)
                            fNewRow = False
                        End If

                        ' Code for the column (padded to specified length)
                        'NPG20090313 Fault 13595
                        ' If bCMGExportFieldCode Then pstrExportString.Append strCMGCode
                        If iCMGExportFieldCodeOrderID = iRow And bCMGExportFieldCode Then
                            ' Insert the delimeter if required.
                            If bUseCSV And Not fNewRow Then pstrExportString.Append(",") ' may need to set fnewrow to false here...
                            pstrExportString.Append(strCMGCode)
                            fNewRow = False
                        End If

                        ' New value
                        ' varNewValue = IIf(IsNull(rsAffectedRecords.Fields(CMGFields.NewValue).Value), vbNullString, rsAffectedRecords.Fields(CMGFields.NewValue).Value)

                        'NPG20071218 Fault 12867
                        Select Case strConvertCase
                            Case "1"
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                varNewValue = IIf(IsDBNull(rsAffectedRecords.Fields(CMGFields.NewValue).value), vbNullString, UCase(rsAffectedRecords.Fields(CMGFields.NewValue).value))
                            Case "2"
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                varNewValue = IIf(IsDBNull(rsAffectedRecords.Fields(CMGFields.NewValue).value), vbNullString, LCase(rsAffectedRecords.Fields(CMGFields.NewValue).value))
                            Case Else
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                varNewValue = IIf(IsDBNull(rsAffectedRecords.Fields(CMGFields.NewValue).value), vbNullString, rsAffectedRecords.Fields(CMGFields.NewValue).value)
                        End Select

                        ' Format the new value
                        If iDataType = 11 Then
                            strExportColumn = VB6.Format(varNewValue, mstrExportDateFormat)
                        Else
                            strExportColumn = varNewValue
                        End If



                        'NPG20090313 Fault 13595

                        If iCMGExportFileCodeOrderID + iCMGEXportRecordIDOrderID + iCMGExportFieldCodeOrderID + iCMGExportOutputColumnOrderID + iCMGExportLastChangeDateOrderID = 0 Then
                            ' The CMG layout hasn't been saved in v3.7 format, use existing values (including reverseoutput option)

                            ' ePayFact or normal CMG output
                            If bCMGReverseOutput Then

                                ' Last changed date
                                If bCMGExportLastChangeDate Then
                                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                    dLastChangeDate = IIf(IsDBNull(rsAffectedRecords.Fields(CMGFields.DateTime).value), #12/31/9999#, rsAffectedRecords.Fields(CMGFields.DateTime).value)
                                    pstrExportString.Append(IIf(bUseCSV, "," & VB6.Format(dLastChangeDate, mstrExportDateFormat), SetStringLength(VB6.Format(dLastChangeDate, mstrExportDateFormat), iCMGExportLastChangeDateSize)))
                                End If

                                ' Add the output column (padded/trimmed to specified width)
                                If bUseCSV Then
                                    pstrExportString.Append(IIf(bUseCSV, ",", vbNullString) & Trim(strExportColumn))
                                Else
                                    pstrExportString.Append(SetStringLength(strExportColumn, iCMGExportOutputColumnSize))
                                End If

                            Else

                                ' Add the output column (padded/trimmed to specified width)
                                If bUseCSV Then
                                    pstrExportString.Append("," & Trim(strExportColumn))
                                Else
                                    pstrExportString.Append(SetStringLength(strExportColumn, iCMGExportOutputColumnSize))
                                End If

                                ' Last changed date
                                If bCMGExportLastChangeDate Then
                                    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                    dLastChangeDate = IIf(IsDBNull(rsAffectedRecords.Fields(CMGFields.DateTime).value), #12/31/9999#, rsAffectedRecords.Fields(CMGFields.DateTime).value)
                                    pstrExportString.Append(IIf(bUseCSV, "," & VB6.Format(dLastChangeDate, mstrExportDateFormat), SetStringLength(VB6.Format(dLastChangeDate, mstrExportDateFormat), iCMGExportLastChangeDateSize)))
                                End If
                            End If
                        Else
                            ' Using new CMG layout ordering...
                            If iCMGExportOutputColumnOrderID = iRow Then
                                ' If iRow < 4 And bUseCSV Then pstrExportString.Append ","
                                If bUseCSV Then
                                    ' Insert the delimeter if required.
                                    If Not fNewRow Then pstrExportString.Append(",") ' may need to set fnewrow to false here...
                                    pstrExportString.Append(Trim(strExportColumn))
                                    fNewRow = False
                                Else
                                    pstrExportString.Append(SetStringLength(strExportColumn, iCMGExportOutputColumnSize))
                                    fNewRow = False
                                End If
                            End If

                            If iCMGExportLastChangeDateOrderID = iRow And bCMGExportLastChangeDate Then
                                ' If iRow < 4 And bUseCSV Then pstrExportString.Append ","
                                'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                                dLastChangeDate = IIf(IsDBNull(rsAffectedRecords.Fields(CMGFields.DateTime).value), #12/31/9999#, rsAffectedRecords.Fields(CMGFields.DateTime).value)
                                ' Insert the delimeter if required.
                                If bUseCSV And Not fNewRow Then pstrExportString.Append(",") ' may need to set fnewrow to false here...
                                pstrExportString.Append(IIf(bUseCSV, VB6.Format(dLastChangeDate, mstrExportDateFormat), SetStringLength(VB6.Format(dLastChangeDate, mstrExportDateFormat), iCMGExportLastChangeDateSize)))
                                fNewRow = False
                            End If

                        End If

                    Next iRow

                    mlngExportRecordCount = mlngExportRecordCount + 1

                End If

                ' Scroll to next record
                rsAffectedRecords.MoveNext()

                ' CMG Special - Line feed after after string
                If Not rsAffectedRecords.EOF And pstrExportString.Length <> 0 Then
                    'NPG20080731 Fault 13305
                    'pstrExportString.Append vbNewLine
                    fNewLine = True
                    fNewRow = True
                End If

            Loop

            ' Process the file name
            mstrOutputFileName = ReplaceFormatExpressions(mstrOutputFileName, 1, mlngExportRecordCount)

            ' Output all the columns for this record ID
            If pstrExportString.Length <> 0 Then
                ' Open the file if not already opened
                If Not bExportFileOpened Then

                    ' If filename specified already exists then delete it first.
                    'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    If Len(Dir(mstrOutputFileName)) > 0 Then

                        'NPG20080421 Fault 13113
                        '          If mlngOutputSaveExisting = 3 Then  ' Append To File
                        '            bAppended = True
                        '          Else
                        '            ' NPG20080218 Fault 12778
                        '            If mlngOutputSaveExisting = 2 Then    ' Add Sequential Number to Filename
                        '              mstrOutputFilename = GetSequentialNumberedFile(mstrOutputFilename)
                        '              bAppended = False
                        '            Else
                        '              bAppended = False
                        '              Kill mstrOutputFilename
                        '            End If
                        '          End If

                        Select Case mlngOutputSaveExisting
                            Case 0 ' Overwrite
                                bAppended = False
                                Kill(mstrOutputFileName)

                            Case 1 ' Do not overwrite (fail)
                                mstrErrorString = "File already exists."
                                ExportData_CMGfile = False
                                Exit Function

                            Case 2 ' Add sequential number to filename
                                mstrOutputFileName = GetSequentialNumberedFile(mstrOutputFileName)
                                bAppended = False

                            Case 3 ' Append to Existing File
                                bAppended = True

                        End Select

                    Else
                        bAppended = False
                    End If

                    ' Open file for output.
                    pintFileNo = FreeFile()

                    If bAppended Then
                        FileOpen(pintFileNo, mstrOutputFileName, OpenMode.Append)
                    Else
                        FileOpen(pintFileNo, mstrOutputFileName, OpenMode.Output)
                    End If

                    bExportFileOpened = True
                End If

                PrintLine(pintFileNo, pstrExportString.ToString_Renamed)
            End If

            ' Move to next record within the selected filter
            prstReadyToExport.MoveNext()
            lngRecordCount = lngRecordCount + 1

            rsAffectedRecords.Close()

        Loop

        prstReadyToExport.Close()

        ' Close the final output file
        FileClose(pintFileNo)

        ' If user cancels the export, abort
        If gobjProgress.Cancelled Then
            mblnUserCancelled = True
            ExportData_CMGfile = False
            Exit Function
        Else
            gobjProgress.Bar1Caption = "Updating Audit Records..."
        End If

        ' Update which records have been exported
        ReDim Preserve astrRecordIDs(lngRecordCount - 1)

        strColumnIDs = Join(astrColumnIDs, ",")
        'strRecordIDs = Join(astrRecordIDs, ",")

        For iLoop = 0 To UBound(astrBulkRecordIDs)
            strRecordIDs = astrBulkRecordIDs(iLoop)

            sSQL = "Update asrsysAuditTrail Set CMGExportDate = Convert(smalldatetime,getdate()) " & "Where RecordID IN (" & strRecordIDs & ") AND ColumnID IN (" & strColumnIDs & ")"
            gADOCon.Execute(sSQL,  , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Next iLoop

        ' If we are to auto commit, do it here
        If mbUpdateAuditLog = True Then
            datGeneral.CMGCommit()
        End If

        'Drop the export table as we dont need it now.
        mclsData.ExecuteSql(("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & "DROP TABLE " & mstrTempTableName))

        'UPGRADE_NOTE: Object prstReadyToExport may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        prstReadyToExport = Nothing
        ExportData_CMGfile = True

        ' Clear up the connection
        'UPGRADE_NOTE: Object rsAffectedRecords may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsAffectedRecords = Nothing

        Exit Function

ExportData_CMGfile_ERROR:

        Select Case Err.Number
            Case 70 ' file creation error / sharing violation
                mstrErrorString = "Error creating file '" & mstrOutputFileName & "'. File may already be in use." & vbCrLf & "(" & Err.Description & ")"
            Case 76 ' path not found error
                mlngExportRecordCount = 0
                mstrErrorString = "Error whilst exporting to CMG file." & vbCrLf & "(" & Err.Description & ")"
            Case Else
                mstrErrorString = "Error whilst exporting to CMG file." & vbCrLf & "(" & Err.Description & ")"
                FileClose(pintFileNo)
        End Select

        ExportData_CMGfile = False

    End Function

    Private Function GetUniqueHeading(ByRef intIndex As Short) As Object

        Dim strTempHeading As String
        Dim lngCount As Integer
        Dim lngLoop As Integer
        Dim blnFound As Boolean

        lngCount = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, intIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strTempHeading = mvarColDetails(13, intIndex)
        If strTempHeading = vbNullString Then
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(4, intIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strTempHeading = mvarColDetails(4, intIndex)
        End If

        Do

            blnFound = False
            For lngLoop = 1 To intIndex - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If LCase(mvarColDetails(14, lngLoop)) = LCase(strTempHeading) Then
                    lngCount = lngCount + 1
                    blnFound = True
                End If
            Next

            If blnFound Then
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strTempHeading = mvarColDetails(13, intIndex) & CStr(lngCount)
            End If

        Loop While blnFound

        'UPGRADE_WARNING: Couldn't resolve default property of object GetUniqueHeading. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetUniqueHeading = strTempHeading

    End Function

    Public Function GetSequentialNumberedFile(ByVal strFileName As String) As String

        Dim lngFound As Integer
        Dim lngCount As Integer

        lngCount = 2
        lngFound = InStrRev(strFileName, ".")
        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        Do While Dir(Left(strFileName, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFileName, lngFound)) <> vbNullString
            lngCount = lngCount + 1
        Loop
        GetSequentialNumberedFile = Left(strFileName, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFileName, lngFound)

    End Function

    Public Function InsertNumberIntoFilename(ByVal strFileName As String, ByRef intNumber As Short) As String

        Dim lngFound As Integer
        Dim lngCount As Integer

        lngCount = 2
        lngFound = InStrRev(strFileName, ".")

        InsertNumberIntoFilename = Left(strFileName, lngFound - 1) & "(" & CStr(intNumber) & ")" & Mid(strFileName, lngFound)

    End Function


    Private Function ReadDataIntoArray() As Boolean

        'Dim pintFileNo As Integer
        Dim pstrExportString As DataMgr.clsStringBuilder = New DataMgr.clsStringBuilder
        'Dim pstrDateString As String
        Dim prstReadyToExport As ADODB.Recordset
        Dim pintLoop As Short
        Dim lngRecordNumber As Integer
        Dim bAppended As Boolean
        Dim bForceHeader As Boolean
        Dim tmpDec As Integer
        Dim tmpLen As Integer
        Dim objExpr As clsExprExpression
        Dim strTemp As String

        Dim lngCol As Integer
        Dim lngRow As Integer
        Dim lngMaxCols As Integer
        Dim lngCount As Integer
        Dim lngExportRows As Integer
        Dim iStartRow As Short

        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Len(Dir(mstrOutputFileName)) > 0 Then
            bAppended = (mlngOutputSaveExisting = 3) 'Check if we are appending to file
        Else
            bAppended = False
        End If

        On Error GoTo ReadDataIntoArray_ERROR


        ' Open the export table as a recordset.
        'Set prstReadyToExport = mclsGeneral.GetRecords("SELECT * FROM [" & mstrTempTableName & "]")
        'Set prstReadyToExport = mclsData.OpenRecordset("SELECT * FROM [" & mstrTempTableName & "]", adOpenForwardOnly, adLockReadOnly)
        prstReadyToExport = mclsData.OpenTableDirect("[" & mstrTempTableName & "]", ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If (prstReadyToExport.BOF And prstReadyToExport.EOF) Then
            mstrErrorString = "No records to export."
            mclsData.ExecuteSql(("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & "DROP TABLE " & mstrTempTableName))
            ReadDataIntoArray = False
            bForceHeader = mbForceHeader
            mbNoRecords = True
        Else
            bForceHeader = False
            mbNoRecords = False
        End If


        'Work out how many columns the array requires...
        lngCount = -1
        lngMaxCols = -1
        For pintLoop = 0 To prstReadyToExport.Fields.Count - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If mvarColDetails(0, pintLoop + 1) = "R" Then 'Carriage Return
                lngCount = 0
            Else
                lngCount = lngCount + 1
            End If
            If lngMaxCols < lngCount Then
                lngMaxCols = lngCount
            End If
        Next



        mblnHeader = False
        ' Omit the header if we are appending to file (if specified)
        If Not (bAppended And mbOmitHeader) Or bForceHeader Then

            If mintExportHeader = 2 Or mintExportHeader = 3 Then 'Custom Heading
                mblnHeader = True
                ReDim mstrArrayHeader(lngMaxCols, 0)
                mstrArrayHeader(0, 0) = ReplaceFormatExpressions(mstrExportHeaderText, 1, mlngExportRecordCount)
                iStartRow = 1
            End If

            ' Column Names
            If mintExportHeader = 1 Or mintExportHeader = 3 Then
                mblnHeader = True
                lngCol = 0
                lngRow = iStartRow
                ReDim Preserve mstrArrayHeader(lngMaxCols, iStartRow)
                For pintLoop = 1 To prstReadyToExport.Fields.Count
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If mvarColDetails(0, pintLoop) = "R" Then 'Carriage Return
                        lngCol = -1
                        lngRow = lngRow + 1
                        ReDim Preserve mstrArrayHeader(lngMaxCols, lngRow)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mstrArrayHeader(lngCol, lngRow) = mvarColDetails(13, pintLoop)
                        If mlngOutputFormat = modEnums.OutputFormats.fmtFixedLengthFile Then
                            SetLength(mstrArrayHeader(lngCol, lngRow), mvarColDetails(6, pintLoop))
                        End If
                    End If
                    lngCol = lngCol + 1
                Next
            End If

        End If

        ' If we only need to output the header bomb out here
        lngRow = -1
        mblnFooter = False
        If Not mbNoRecords Then
            'Close #pintFileNo
            'Exit Function

            ReDim mstrArrayData(lngMaxCols, 0)

            ' Loop through the export table and print stuff to file.
            lngRecordNumber = 0
            Do While Not prstReadyToExport.EOF

                ' If user cancels the export, abort
                If (lngRow Mod 100) = 0 Then
                    If gobjProgress.Cancelled Then
                        mblnUserCancelled = True
                        'Close #pintFileNo
                        ReadDataIntoArray = False
                        Exit Function
                    End If
                End If

                lngCol = -1
                lngRow = lngRow + 1
                If lngRow > UBound(mstrArrayData, 2) Then ReDim Preserve mstrArrayData(lngMaxCols, lngRow + 1000)

                If mbLoggingExportSuccess Then
                    pstrExportString.TheString = vbNullString
                End If

                ' Loop through the fields in each record, adding them and the delimiter to the export string.
                lngRecordNumber = lngRecordNumber + 1
                For pintLoop = 0 To prstReadyToExport.Fields.Count - 1

                    lngCol = lngCol + 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If mvarColDetails(0, pintLoop + 1) = "N" Then
                        'Record Number
                        mstrArrayData(lngCol, lngRow) = CStr(lngRecordNumber)

                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ElseIf mvarColDetails(0, pintLoop + 1) = "R" Then
                        'Carriage Return
                        lngCol = -1
                        lngRow = lngRow + 1
                        If lngRow > UBound(mstrArrayData, 2) Then ReDim Preserve mstrArrayData(lngMaxCols, lngRow + 1000)

                        'MH20040402 Fault 8434
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ElseIf mvarColDetails(0, pintLoop + 1) = "F" Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mstrArrayData(lngCol, lngRow) = Space(mvarColDetails(6, pintLoop + 1))

                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(7, pintLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ElseIf mvarColDetails(7, pintLoop + 1) = True Then
                        mstrArrayData(lngCol, lngRow) = VB6.Format(prstReadyToExport.Fields(pintLoop).value, mstrExportDateFormat)

                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, pintLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ElseIf mvarColDetails(12, pintLoop + 1) Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(6, pintLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        tmpLen = mvarColDetails(6, pintLoop + 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(11, pintLoop + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        tmpDec = mvarColDetails(11, pintLoop + 1)
                        'TM19032004 Fault 8052 - DO NOT format Null values to zero. Set to empty string instead.
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        If IsDBNull(prstReadyToExport.Fields(pintLoop).value) Then
                            strTemp = vbNullString
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object FormatNumeric(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strTemp = FormatNumeric((prstReadyToExport.Fields(pintLoop).value), tmpLen, tmpDec)
                        End If

                        mstrArrayData(lngCol, lngRow) = strTemp

                    Else
                        'MH20030401 Fault 5225
                        'mstrArrayData(lngCol, lngRow) = Left(prstReadyToExport.Fields(pintLoop), mvarColDetails(6, pintLoop + 1))
                        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
                        If IsDBNull(prstReadyToExport.Fields(pintLoop).value) Then
                            mstrArrayData(lngCol, lngRow) = vbNullString
                        Else
                            mstrArrayData(lngCol, lngRow) = prstReadyToExport.Fields(pintLoop).value
                        End If
                    End If


                    'NPG20080718 Fault 13276
                    Select Case mvarColDetails(16, lngCol + 1)
                        Case 1 ' Upper Case
                            mstrArrayData(lngCol, lngRow) = UCase(mstrArrayData(lngCol, lngRow))
                        Case 2 ' Lower Case
                            mstrArrayData(lngCol, lngRow) = LCase(mstrArrayData(lngCol, lngRow))
                    End Select

                    If mblnStripDelimiter Then
                        mstrArrayData(lngCol, lngRow) = Replace(mstrArrayData(lngCol, lngRow), mstrExportActualDelimiter, "")
                    End If

                    If lngCol >= 0 Then
                        'Add spaces for fixed length or trim is data is too long...
                        SetLength(mstrArrayData(lngCol, lngRow), mvarColDetails(6, pintLoop + 1))

                        If mbLoggingExportSuccess Then
                            pstrExportString.Append(IIf((pstrExportString.Length <> 0), mstrExportDelimiter, vbNullString) & mstrArrayData(lngCol, lngRow))
                        End If
                    End If

                Next pintLoop

                ' Print the stuff to the file.
                'Print #pintFileNo, pstrExportString


                'JDM - 13/12/01 - Fault 3280 - Log successful records
                If mbLoggingExportSuccess Then
                    gobjEventLog.AddDetailEntry(pstrExportString.ToString_Renamed & vbCrLf & "Exported successfully")
                End If

                mlSuccessfulRecords = mlSuccessfulRecords + 1

                prstReadyToExport.MoveNext()

            Loop

            ReDim Preserve mstrArrayData(lngMaxCols, lngRow)

            If mintExportFooter = 2 Then 'Custom Footing
                mblnFooter = True
                ReDim mstrArrayFooter(lngMaxCols, 0)
                mstrArrayFooter(0, 0) = ReplaceFormatExpressions(mstrExportFooterText, 1, mlngExportRecordCount)
            End If

            If mintExportFooter = 1 Then 'Column Names
                mblnFooter = True
                lngCol = 0
                lngRow = 0
                ReDim mstrArrayFooter(lngMaxCols, 0)
                For pintLoop = 1 To prstReadyToExport.Fields.Count
                    'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(0, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If mvarColDetails(0, pintLoop) = "R" Then 'Carriage Return
                        lngCol = -1
                        lngRow = lngRow + 1
                        ReDim Preserve mstrArrayFooter(lngMaxCols, lngRow)
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(13, pintLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mstrArrayFooter(lngCol, lngRow) = mvarColDetails(13, pintLoop)
                        If mlngOutputFormat = modEnums.OutputFormats.fmtFixedLengthFile Then
                            SetLength(mstrArrayFooter(lngCol, lngRow), mvarColDetails(6, pintLoop))
                        End If
                    End If
                    lngCol = lngCol + 1
                Next
            End If

            ' Close the final output file
            'Close #pintFileNo

        End If

        'Debug.Print "end" & Now

        'Drop the export table as we dont need it now.
        'UPGRADE_NOTE: Object prstReadyToExport may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        prstReadyToExport = Nothing
        mclsData.ExecuteSql(("IF EXISTS(SELECT * FROM sysobjects WHERE name = '" & mstrTempTableName & "') " & "DROP TABLE " & mstrTempTableName))
        ReadDataIntoArray = True
        Exit Function

ReadDataIntoArray_ERROR:

        'Select Case Err.Number
        'Case 70 ' file creation error / sharing violation
        '  mstrErrorString = "Error creating file '" & mstrOutputFilename & "'. File may already be in use." & vbCrLf & "(" & Err.Description & ")"
        'Case 76 ' path not found error
        '  mlngExportRecordCount = 0
        '  mstrErrorString = "Error whilst exporting to delimited file." & vbCrLf & "(" & Err.Description & ")"
        'Case Else
        mstrErrorString = "Error whilst exporting data." & vbCrLf & "(" & Err.Description & ")"
        '  Close #pintFileNo
        'End Select

        ReadDataIntoArray = False

    End Function

    Private Function SendArrayToOutputOptions() As Object

        Dim objOutput As clsOutputRun
        Dim objColumn As clsColumn
        Dim lngLoop As Integer
        Dim intBlockLoop As Short
        Dim intBlockTotal As Short
        Dim sExtension As String
        Dim sFileName As String
        Dim aryOutputData() As String

        objOutput = New clsOutputRun

        ' How many files required for this export?
        If mbSplitFile Then
            intBlockTotal = Floor(UBound(mstrArrayData, 2) / mlngSplitFileSize) + 1

            ' If split into blocks ensure that there is some logic for the filename
            If InStr(mstrOutputFileName, "{BLOCKNUMBER}") = 0 Then
                sExtension = "." & Mid(mstrOutputFileName, InStrRev(mstrOutputFileName, ".") + 1, Len(mstrOutputFileName))
                mstrOutputFileName = Replace(mstrOutputFileName, sExtension, "_{BLOCKNUMBER}" & sExtension)
            End If

        Else
            intBlockTotal = 1
        End If

        For intBlockLoop = 1 To intBlockTotal

            ' Get a block of exporting data
            If mbSplitFile And mlngSplitFileSize > 0 Then
                aryOutputData = OutputArrayBlock(intBlockLoop)

                If mblnHeader Then
                    mstrArrayHeader(0, 0) = ReplaceFormatExpressions(mstrExportHeaderText, intBlockLoop, UBound(aryOutputData, 2) + 1)
                End If

                If mblnFooter Then
                    mstrArrayFooter(0, 0) = ReplaceFormatExpressions(mstrExportFooterText, intBlockLoop, UBound(aryOutputData, 2) + 1)
                End If

            Else
                aryOutputData = VB6.CopyArray(mstrArrayData)
            End If

            ' Process the file name
            sFileName = ReplaceFormatExpressions(mstrOutputFileName, intBlockLoop, UBound(aryOutputData, 2))

            'UPGRADE_WARNING: Couldn't resolve default property of object objOutput.SetOptions(False, mlngOutputFormat, False, False, vbNullString, mblnOutputSave, mlngOutputSaveExisting, mblnOutputEmail, mlngOutputEmailAddr, mstrOutputEmailSubject, mstrOutputEmailAttachAs, sFileName). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If objOutput.SetOptions(False, mlngOutputFormat, False, False, vbNullString, mblnOutputSave, mlngOutputSaveExisting, mblnOutputEmail, mlngOutputEmailAddr, mstrOutputEmailSubject, mstrOutputEmailAttachAs, sFileName) Then

                objOutput.FileDelimiter = mstrExportActualDelimiter
                objOutput.ApplyStyles = False
                objOutput.SizeColumnsIndependently = True

                If objOutput.GetFile Then

                    For lngLoop = 1 To UBound(mvarColDetails, 2)
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'UPGRADE_WARNING: Couldn't resolve default property of object mvarColDetails(12, lngLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        objOutput.AddColumn(CStr(mvarColDetails(13, lngLoop)), IIf(mvarColDetails(12, lngLoop), modEnums.SQLDataType.sqlNumeric, modEnums.SQLDataType.sqlVarChar), CInt(mvarColDetails(11, lngLoop)))
                    Next

                    objOutput.HeaderRows = 0
                    objOutput.AddPage(mstrExportName, mstrExportName)

                    If mblnHeader Then
                        objOutput.DisableDelimiterCheck = True
                        objOutput.DataArray(mstrArrayHeader)
                    End If

                    objOutput.EncloseInQuotes = mblnExportQuotes
                    If Not mblnNoRecords Then
                        objOutput.DisableDelimiterCheck = False
                        objOutput.UpdateProgressPerRow = True
                        objOutput.DataArray(aryOutputData)
                        objOutput.UpdateProgressPerRow = False
                    End If
                    objOutput.EncloseInQuotes = False

                    If mblnFooter Then
                        objOutput.DisableDelimiterCheck = True
                        objOutput.DataArray(mstrArrayFooter)
                    End If

                    If Not gblnBatchMode Then
                        gobjProgress.CloseProgress()
                    End If
                    objOutput.Complete()

                End If

            End If

        Next intBlockLoop

        objOutput.ClearUp()
        mstrErrorString = objOutput.ErrorMessage

        'UPGRADE_NOTE: Object objOutput may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objOutput = Nothing
        'UPGRADE_WARNING: Couldn't resolve default property of object SendArrayToOutputOptions. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        SendArrayToOutputOptions = (mstrErrorString = vbNullString)

    End Function


    Private Sub SetLength(ByRef strInput As String, ByRef lngLength As Object)

        Dim blnFixedLength As Boolean

        blnFixedLength = (mlngOutputFormat = modEnums.OutputFormats.fmtFixedLengthFile)

        'UPGRADE_WARNING: Couldn't resolve default property of object lngLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If blnFixedLength Or lngLength > 0 Then
            If blnFixedLength Then
                'UPGRADE_WARNING: Couldn't resolve default property of object lngLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strInput = strInput & Space(lngLength)
            End If

            ' If they've banged in a very big size ignore it as Left$ can only trim so much...
            'UPGRADE_WARNING: Couldn't resolve default property of object lngLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If lngLength < 999999999 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object lngLength. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strInput = Left(strInput, lngLength)
            End If

        End If

    End Sub

    Private Function GetFileText(ByVal FileName As String) As String
        Dim Handle As Short
        Handle = FreeFile()
        FileOpen(Handle, FileName, OpenMode.Input)
        GetFileText = InputString(Handle, LOF(Handle))
        FileClose(Handle)
    End Function

    ' Extracts just the filename from a path
    Function GetFileNameOnly(ByRef pstrFilePath As String) As String
        Dim astrPath() As String
        astrPath = Split(pstrFilePath, "\")
        GetFileNameOnly = astrPath(UBound(astrPath))
    End Function

    Private Function ReplaceFormatExpressions(ByVal value As String, ByVal blockNumber As Integer, ByVal blockRecordCount As Integer) As String

        ' Replace variables
        value = Replace(value, "{BLOCKNUMBER}", CStr(blockNumber))
        value = Replace(value, "{BLOCKCOUNT}", CStr(blockRecordCount))
        value = Replace(value, "{RECORDCOUNT}", CStr(mlngExportRecordCount))
        value = Replace(value, "{DATETIME}", VB6.Format(mdExportCreateDate, mstrExportDateFormat & "hhnnss"))
        'UPGRADE_WARNING: DateValue has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        value = Replace(value, "{DATE}", VB6.Format(DateValue(CStr(mdExportCreateDate)), mstrExportDateFormat))

        ReplaceFormatExpressions = value

    End Function

    Private Function Floor(ByVal dblValue As Double) As Double
        Dim myDec As Integer
        myDec = InStr(1, CStr(dblValue), ".", CompareMethod.Text)
        If myDec > 0 Then
            Floor = CDbl(Left(CStr(dblValue), myDec))
        Else
            Floor = dblValue
        End If
    End Function

    Private Function OutputArrayBlock(ByVal intBlockNumber As Short) As String()

        On Error GoTo ArrayExceeded

        Dim iStartBlock As Short
        Dim iRowsInBlock As Short
        Dim lngTotalRows As Integer
        Dim iColumnLoop As Short
        Dim j As Short

        lngTotalRows = UBound(mstrArrayData, 2)
        iStartBlock = (intBlockNumber - 1) * mlngSplitFileSize
        iRowsInBlock = Minimum(mlngSplitFileSize, CInt(lngTotalRows - iStartBlock) + 1)
        Dim aryOutput(UBound(mstrArrayData, 1), iRowsInBlock - 1) As String

        For iColumnLoop = 0 To UBound(mstrArrayData, 1)
            For j = 0 To iRowsInBlock - 1
                aryOutput(iColumnLoop, j) = mstrArrayData(iColumnLoop, j + iStartBlock)
            Next j
        Next iColumnLoop

ArrayExceeded:
        OutputArrayBlock = VB6.CopyArray(aryOutput)

    End Function
End Class