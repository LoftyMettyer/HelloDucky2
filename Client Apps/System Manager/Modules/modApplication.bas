Attribute VB_Name = "modApplication"
Option Explicit

Public Function Activate() As Boolean
  On Error GoTo ErrorTrap

  If Not Application.LoggedIn Then
    Login
  End If
  
  ' NPG20090924 Fault HRPro-330
  Dim frmStyle As New frmHiddenStyle
  Load frmStyle

  If IsFormLoaded("frmHiddenStyle") Then
    UnLoad frmStyle
  End If
  
  If Application.LoggedIn Then
    'Create temporary local database
    If CreateTempDb() Then
        
      'Create temparary tables
      If CreateTempTables() Then
      
        ' Load defintions into system framework (ultimately will replace the above temp tables, but this has to be done piecemeal)
        PopulateMetaData
      
        CreateQueryDefs
      
        ActivateModules
        
        'Load system manager form
        frmSysMgr.SetBackground (True)
        frmSysMgr.Show
        Activate = True
      End If
    End If
  End If
  
  Exit Function

ErrorTrap:
  Activate = False
  Err = False
  
End Function

Public Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim frm As Form


    For Each frm In Forms
        If LCase(frm.Name) = LCase(FormName) Then IsFormLoaded = True
    Next frm
End Function

Public Function CreateTempDb() As Boolean
  On Error GoTo ErrorTrap
  
  Dim strTempPath As String
  Dim bCreateNew As Boolean
  
  bCreateNew = True
  
  'Get TEMP environment variable setting
  strTempPath = Environ("TEMP")
  
  'Generate the Temp Database Name
  gsTempDatabaseName = "AsrTemp_" & gsDatabaseName & ".mdb"
  
  ' This recovery mode is a development/diagnostic tool - not designed to be used customers
  If gbAttemptRecovery Then
    If Dir(strTempPath & "\" & gsTempDatabaseName) <> vbNullString Then
      Set daoWS = CreateWorkspace("Temp", "Admin", vbNullString, dbUseJet)
      Set daoDb = daoWS.OpenDatabase(strTempPath & "\" & gsTempDatabaseName)
      bCreateNew = False
    Else
      gbAttemptRecovery = False
    End If
  End If
  
  If bCreateNew Then
    'Make sure temp ASR database does not already exist
    If Dir(strTempPath & "\" & gsTempDatabaseName) <> vbNullString Then
      Kill strTempPath & "\" & gsTempDatabaseName
    End If
  
    'Create new workspace for temporary database
    Set daoWS = CreateWorkspace("Temp", "Admin", vbNullString, dbUseJet)
    
    'Create new temp ASR database
    Set daoDb = daoWS.CreateDatabase(strTempPath & "\" & gsTempDatabaseName, dbLangGeneral)
  
    ' Allow access via external DLLs
    daoDb.Close
    Set daoDb = daoWS.OpenDatabase(strTempPath & "\" & gsTempDatabaseName, False)
  
  End If
  
  CreateTempDb = True
  Exit Function

ErrorTrap:

  Select Case Err.Number
  
    'MH20010702 Seems to be error 75 on Win98 PCs...
    'Case 70:
    Case 70, 75:
      
        MsgBox "An instance of the System Manager is already running on this machine" & vbCrLf & _
               "using the '" & gsDatabaseName & "' database." & vbCrLf & vbCrLf & _
               "Each instance of the System Manager must be accessing a unique database.", _
               vbCritical + vbOKOnly, App.Title

    Case Else
      MsgBox "Error creating temp DB." & vbCrLf & Err.Description, vbCritical + vbOKOnly, App.Title

  End Select
  
  CreateTempDb = False
  
End Function

Public Sub PopulateMetaData()

  On Error GoTo ErrorTrap

'gobjMobileDefs.Connection = gADOCon
'gobjMobileDefs.Populate

  Exit Sub

ErrorTrap:

  gobjProgress.Visible = False
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Name
  Err = False

End Sub

Public Function CreateTempTables() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim sSource As String
  Dim tbdTemp As DAO.TableDef
  Dim idxTemp As DAO.Index


  Application.ChangedTableName = False
  Application.ChangedViewName = False
  Application.ChangedColumnName = False

  ' Get the source for table creation queries.
  
  'MH20010704 When doing the SELECT statement pulling data from SQL into Access another
  'connection is created.  Give this connection a different app name otherwise it looks
  'like the Sys Mgr user is logged on twice !
  
  If Not gbAttemptRecovery Then
    
    If gbUseWindowsAuthentication Then
      sSource = "'' [ODBC;DRIVER=SQL Server;DSN=" & gstrWindowsAuthentication_DNSName & ";SERVER=" & gsServerName _
          & ";APP=Access Database;" _
          & ";DATABASE=" & gsDatabaseName & "]"
    Else
      sSource = "'' [ODBC;DRIVER=SQL Server;SERVER=" & gsServerName _
          & ";UID=" & gsUserName & ";PWD=" & gsPassword & ";APP=Access Database;" _
          & ";DATABASE=" & gsDatabaseName & "]"
    End If
    
    sSQL = "SELECT ASRSysTables.*," & _
      " ASRSysTables.TableName as OriginalTableName, " & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted, 0 AS copyDataTableID, FALSE AS IsCopy " & _
      " , 0 AS copySecurityTableID, '' As CopySecurityTableName, 0 As GrantRead, 0 AS GrantNew, 0 As GrantEdit, 0 As GrantDelete, 0 AS PermissionsPrompted " & _
      " INTO tmpTables FROM ASRSysTables IN " & sSource
    daoDb.Execute sSQL
  
    
    ' Create the local SummaryFields table.
    sSQL = "SELECT ASRSysSummaryFields.*" & _
      " INTO tmpSummary FROM ASRSysSummaryFields IN " & sSource
    daoDb.Execute sSQL
  
   
    sSQL = "SELECT ASRSysColumns.*," & _
      " ColumnName as OriginalColumnName, DataType as OriginalDataType," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpColumns FROM ASRSysColumns IN " & sSource
    daoDb.Execute sSQL
  
  
    'TM20011005 Fault 2192 - 4
    'Added the Changed, New, Deleted flags to the Diary Links temp table.
    ' Create the local Diary table.
    sSQL = "SELECT ASRSysDiaryLinks.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpDiary FROM ASRSysDiaryLinks IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Column Control Values table.
    sSQL = "SELECT ASRSysColumnControlValues.* INTO tmpControlValues" & _
      " FROM ASRSysColumnControlValues IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Relation table.
    sSQL = "SELECT ASRSysRelations.* INTO tmpRelations" & _
      " FROM ASRSysRelations IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Table Validations table.
    sSQL = "SELECT ASRSysTableValidations.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpTableValidations FROM ASRSysTableValidations IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local History Screens table.
    sSQL = "SELECT ASRSysHistoryScreens.* INTO tmpHistoryScreens" & _
      " FROM ASRSysHistoryScreens IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Screens table.
    sSQL = "SELECT ASRSysScreens.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpScreens FROM ASRSysScreens IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Screen Page Captions table.
    sSQL = "SELECT ASRSysPageCaptions.* INTO tmpPageCaptions" & _
      " FROM ASRSysPageCaptions IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Controls table.
    sSQL = "SELECT ASRSysControls.* INTO tmpControls" & _
      " FROM ASRSysControls IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Pictures table.
    sSQL = "SELECT ASRSysPictures.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpPictures FROM ASRSysPictures IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Orders table.
    sSQL = "SELECT ASRSysOrders.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpOrders FROM ASRSysOrders IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local OrderItems table.
    sSQL = "SELECT ASRSysOrderItems.*" & _
      " INTO tmpOrderItems FROM ASRSysOrderItems IN " & sSource
    daoDb.Execute sSQL
       
    
    ' Create the local Email Address table.
    sSQL = "SELECT ASRSysEmailAddress.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpEmailAddresses FROM ASRSysEmailAddress IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Email Links table.
    sSQL = "SELECT ASRSysEmailLinks.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpEmailLinks FROM ASRSysEmailLinks IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Email Links Attachments table.
    'sSQL = "SELECT ASRSysEmailLinksAttachments.*" & _
      " INTO tmpEmailLinksAttachments FROM ASRSysEmailLinksAttachments IN " & sSource
    'daoDb.Execute sSQL
    
    
    sSQL = "SELECT ASRSysEmailLinksColumns.*" & _
      " INTO tmpEmailLinksColumns FROM ASRSysEmailLinksColumns IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Email Links Recipients table.
    sSQL = "SELECT ASRSysEmailLinksRecipients.*" & _
      " INTO tmpEmailLinksRecipients FROM ASRSysEmailLinksRecipients IN " & sSource
    daoDb.Execute sSQL
    'MH20000727
    
    
    'MH20090520
    sSQL = "SELECT ASRSysLinkContent.*" & _
      " INTO tmpLinkContent FROM ASRSysLinkContent IN " & sSource
    daoDb.Execute sSQL
    
    
    
    'MH20040301 Outlook Calendar Interface
    sSQL = "SELECT ASRSysOutlookFolders.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpOutlookFolders FROM ASRSysOutlookFolders IN " & sSource
    daoDb.Execute sSQL
  
    sSQL = "SELECT ASRSysOutlookLinks.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpOutlookLinks FROM ASRSysOutlookLinks IN " & sSource
    daoDb.Execute sSQL
  
    sSQL = "SELECT ASRSysOutlookLinksColumns.*" & _
      " INTO tmpOutlookLinksColumns FROM ASRSysOutlookLinksColumns IN " & sSource
    daoDb.Execute sSQL
  
    sSQL = "SELECT ASRSysOutlookLinksDestinations.*" & _
      " INTO tmpOutlookLinksDestinations FROM ASRSysOutlookLinksDestinations IN " & sSource
    daoDb.Execute sSQL
  
  
  
    ' Create the local Expressions table.
    sSQL = "SELECT ASRSysExpressions.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted, now AS lastSave" & _
      " INTO tmpExpressions FROM ASRSysExpressions IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Expression Components table.
    sSQL = "SELECT ASRSysExprComponents.*" & _
      " INTO tmpComponents FROM ASRSysExprComponents IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local View table.
    sSQL = "SELECT ASRSysViews.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted," & _
      " ViewName as OriginalViewName," & _
      " 0 As GrantRead, 0 AS GrantNew, 0 As GrantEdit, 0 As GrantDelete " & _
      " INTO tmpViews FROM ASRSysViews IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local View column table.
    sSQL = "SELECT ASRSysViewColumns.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpViewColumns FROM ASRSysViewColumns IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local View screen table.
    sSQL = "SELECT ASRSysViewScreens.*," & _
      " FALSE AS new, FALSE AS deleted" & _
      " INTO tmpViewScreens FROM ASRSysViewScreens IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Module Setup table.
    sSQL = "SELECT ASRSysModuleSetup.*" & _
      " INTO tmpModuleSetup FROM ASRSysModuleSetup IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Module Related Columns table.
    sSQL = "SELECT ASRSysModuleRelatedColumns.*" & _
      " INTO tmpModuleRelatedColumns FROM ASRSysModuleRelatedColumns IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Self-service Intranet Links table.
    sSQL = "SELECT ASRSysSSIntranetLinks.*" & _
      " INTO tmpSSIntranetLinks FROM ASRSysSSIntranetLinks IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Self-service Intranet Hidden Groups table.
    sSQL = "SELECT ASRSysSSIHiddenGroups.*" & _
      " INTO tmpSSIHiddenGroups FROM ASRSysSSIHiddenGroups IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Self-service Intranet Views table.
    sSQL = "SELECT ASRSysSSIViews.*" & _
      " INTO tmpSSIViews FROM ASRSysSSIViews IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Payroll Field Definition table.
    sSQL = "SELECT ASRSysAccordTransferFieldDefinitions.*" & _
      " INTO tmpAccordTransferFieldDefinitions FROM ASRSysAccordTransferFieldDefinitions IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Payroll Mapping table.
    sSQL = "SELECT ASRSysAccordTransferFieldMappings.*" & _
      " INTO tmpAccordTransferFieldMappings FROM ASRSysAccordTransferFieldMappings IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Payroll Transfer Type table.
    sSQL = "SELECT ASRSysAccordTransferTypes.*" & _
      " INTO tmpAccordTransferTypes FROM ASRSysAccordTransferTypes IN " & sSource
    daoDb.Execute sSQL
  
'      ' Create the local Fusion Field Definition table.
'    sSQL = "SELECT ASRSysFusionFieldDefinitions.*" & _
'      " INTO tmpFusionFieldDefinitions FROM ASRSysFusionFieldDefinitions IN " & sSource
'    daoDb.Execute sSQL
'
'    ' Create the local Fusion Mapping table.
'    sSQL = "SELECT ASRSysFusionFieldMappings.*" & _
'      " INTO tmpFusionFieldMappings FROM ASRSysFusionFieldMappings IN " & sSource
'    daoDb.Execute sSQL
'
'    ' Create the local Fusion Transfer Type table.
'    sSQL = "SELECT ASRSysFusionTypes.*" & _
'      " INTO tmpFusionTypes FROM ASRSysFusionTypes IN " & sSource
'    daoDb.Execute sSQL
    
    'TM20020211 Fault 3487
    ' Create the local mail merge table.
    sSQL = "SELECT ASRSysMailMergeName.* " & _
          " INTO tmpMailMerge FROM ASRSysMailMergeName IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Workflows table.
    sSQL = "SELECT ASRSysWorkflows.*," & _
      " FALSE AS changed, FALSE AS perge, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpWorkflows FROM ASRSysWorkflows IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Workflow Elements table.
    sSQL = "SELECT ASRSysWorkflowElements.*" & _
      " INTO tmpWorkflowElements FROM ASRSysWorkflowElements IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Workflow Links table.
    sSQL = "SELECT ASRSysWorkflowLinks.*" & _
      " INTO tmpWorkflowLinks FROM ASRSysWorkflowLinks IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Workflow Element Items table.
    sSQL = "SELECT ASRSysWorkflowElementItems.*" & _
      " INTO tmpWorkflowElementItems FROM ASRSysWorkflowElementItems IN " & sSource
    daoDb.Execute sSQL
    
    ' Create the local Workflow Element Item Values table.
    sSQL = "SELECT ASRSysWorkflowElementItemValues.*" & _
      " INTO tmpWorkflowElementItemValues FROM ASRSysWorkflowElementItemValues IN " & sSource
    daoDb.Execute sSQL

    ' Create the local Workflow Element Columns table.
    sSQL = "SELECT ASRSysWorkflowElementColumns.*" & _
      " INTO tmpWorkflowElementColumns FROM ASRSysWorkflowElementColumns IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Workflow Element Validations table.
    sSQL = "SELECT ASRSysWorkflowElementValidations.*" & _
      " INTO tmpWorkflowElementValidations FROM ASRSysWorkflowElementValidations IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Workflow Triggered Links table.
    sSQL = "SELECT ASRSysWorkflowTriggeredLinks.*," & _
      " FALSE AS changed, FALSE AS new, FALSE AS deleted" & _
      " INTO tmpWorkflowTriggeredLinks FROM ASRSysWorkflowTriggeredLinks IN " & sSource
    daoDb.Execute sSQL
  
    ' Create the local Workflow Triggered Link Columns table.
    sSQL = "SELECT ASRSysWorkflowTriggeredLinkColumns.*" & _
      " INTO tmpWorkflowTriggeredLinkColumns FROM ASRSysWorkflowTriggeredLinkColumns IN " & sSource
    daoDb.Execute sSQL
   
    ' Mobile navigation tables (layouts).
    sSQL = "SELECT tbsys_MobileFormLayout.*" & _
      " INTO tmpmobileformlayout FROM tbsys_MobileFormLayout IN " & sSource
    daoDb.Execute sSQL
  
    ' Mobile navigation tables (elements).
    sSQL = "SELECT tbsys_MobileFormElements.*" & _
      " INTO tmpmobileformelements FROM tbsys_MobileFormElements IN " & sSource
    daoDb.Execute sSQL
  
    ' Local copy of the security groups
    sSQL = "SELECT ASRSysGroups.*" & _
      " INTO tmpGroups FROM ASRSysGroups IN " & sSource
    daoDb.Execute sSQL
   
    ' Mobile group workflows
    sSQL = "SELECT tbsys_mobilegroupworkflows.*" & _
      " INTO tmpmobilegroupworkflows FROM tbsys_mobilegroupworkflows IN " & sSource
    daoDb.Execute sSQL
  
  
    ' Create the indices for the local Tables table.
    Set tbdTemp = daoDb.TableDefs("tmpTables")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("tableID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxName")
      With idxTemp
        .Fields.Append .CreateField("tableName")
        .Fields.Append .CreateField("deleted")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    
    ' Create the indices for the local Table Validations table.
    Set tbdTemp = daoDb.TableDefs("tmpTableValidations")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxValidationID")
      With idxTemp
        .Fields.Append .CreateField("ValidationID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("tableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
     
      .Indexes.Refresh
    End With
    
    
    ' Create the indices for the local Summary Fields table.
    Set tbdTemp = daoDb.TableDefs("tmpSummary")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxID")
      With idxTemp
        .Fields.Append .CreateField("ID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxRealOrder")
      With idxTemp
        .Fields.Append .CreateField("historyTableID")
        .Fields.Append .CreateField("sequence")
        .Fields.Append .CreateField("parentColumnID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Columns table.
    Set tbdTemp = daoDb.TableDefs("tmpColumns")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxColumnID")
      With idxTemp
        .Fields.Append .CreateField("columnID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxName")
      With idxTemp
        .Fields.Append .CreateField("tableID")
        .Fields.Append .CreateField("columnName")
        .Fields.Append .CreateField("deleted")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxLookup")
      With idxTemp
        .Fields.Append .CreateField("lookupTableID")
        .Fields.Append .CreateField("lookupColumnID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxCalcExprID")
      With idxTemp
        .Fields.Append .CreateField("calcExprID")
        .Fields.Append .CreateField("deleted")
        .Unique = False
      End With
      .Indexes.Append idxTemp
          
      ' RJB 12/08/1998
      ' Add a table id index
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("TableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Diary Links table
    Set tbdTemp = daoDb.TableDefs("tmpDiary")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxColumnID")
      With idxTemp
        .Fields.Append .CreateField("columnID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxDiaryLinkID")
      With idxTemp
        .Fields.Append .CreateField("diaryID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Column Control Values table
    Set tbdTemp = daoDb.TableDefs("tmpControlValues")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxColumnID")
      With idxTemp
        .Fields.Append .CreateField("columnID")
        .Fields.Append .CreateField("sequence")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Relations table.
    Set tbdTemp = daoDb.TableDefs("tmpRelations")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxParentID")
      With idxTemp
        .Fields.Append .CreateField("parentID")
        .Fields.Append .CreateField("childID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxChildID")
      With idxTemp
        .Fields.Append .CreateField("childID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxChildParentID")
      With idxTemp
        .Fields.Append .CreateField("childID")
        .Fields.Append .CreateField("parentID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local History Screens table.
    Set tbdTemp = daoDb.TableDefs("tmpHistoryScreens")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxParentHistory")
      With idxTemp
        .Fields.Append .CreateField("parentScreenID")
        .Fields.Append .CreateField("historyScreenID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxHistoryScreenID")
      With idxTemp
        .Fields.Append .CreateField("historyScreenID")
        .Fields.Append .CreateField("order")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Screens table.
    Set tbdTemp = daoDb.TableDefs("tmpScreens")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxScreenID")
      With idxTemp
        .Fields.Append .CreateField("screenID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxName")
      With idxTemp
        .Fields.Append .CreateField("name")
        .Fields.Append .CreateField("deleted")
        ' JPD 19/02/2001 - Changed this index to be non-unique as it caused
        ' errors. Imagine creating a screen called 'screen a', then deleting it.
        ' You can then create another screen called 'screen a' no problem. But if you
        ' delete this screen then there will be two records with the same 'name'
        ' and 'deleted' values. Hence this index cannot be unique.
        '.Unique = True
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("tableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Screen Page Captions table.
    Set tbdTemp = daoDb.TableDefs("tmpPageCaptions")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxScreenPage")
      With idxTemp
        .Fields.Append .CreateField("screenID")
        .Fields.Append .CreateField("pageIndexID")
      End With
      .Indexes.Append idxTemp
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Controls table.
    Set tbdTemp = daoDb.TableDefs("tmpControls")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxScreenID")
      With idxTemp
        .Fields.Append .CreateField("screenID")
        .Fields.Append .CreateField("pageNo")
        .Fields.Append .CreateField("controlLevel")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxTabIndex")
      With idxTemp
        .Fields.Append .CreateField("screenID")
        .Fields.Append .CreateField("pageNo")
        .Fields.Append .CreateField("tabIndex")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Pictures table.
    Set tbdTemp = daoDb.TableDefs("tmpPictures")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxID")
      With idxTemp
        .Fields.Append .CreateField("pictureID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
  
      Set idxTemp = .CreateIndex("idxName")
      With idxTemp
        .Fields.Append .CreateField("name")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Orders table.
    Set tbdTemp = daoDb.TableDefs("tmpOrders")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxID")
      With idxTemp
        .Fields.Append .CreateField("orderID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("tableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxName")
      With idxTemp
        .Fields.Append .CreateField("name")
        .Fields.Append .CreateField("deleted")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Order Items table.
    Set tbdTemp = daoDb.TableDefs("tmpOrderItems")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxOrderID")
      With idxTemp
        .Fields.Append .CreateField("orderID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      .Indexes.Refresh
    End With
    
    
    'MH200000727 Added Email
    ' Create the indices for the local email table.
    Set tbdTemp = daoDb.TableDefs("tmpEmailAddresses")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxID")
      With idxTemp
        .Fields.Append .CreateField("EmailID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("tableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxName")
      With idxTemp
        .Fields.Append .CreateField("name")
        .Fields.Append .CreateField("deleted")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local email table.
    Set tbdTemp = daoDb.TableDefs("tmpEmailLinks")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxID")
      With idxTemp
        .Fields.Append .CreateField("LinkID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("TableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      'Set idxTemp = .CreateIndex("idxColumnID")
      'With idxTemp
      '  .Fields.Append .CreateField("ColumnID")
      '  .Unique = False
      'End With
      '.Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    
    Set tbdTemp = daoDb.TableDefs("tmpEmailLinksColumns")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxLinkID")
      With idxTemp
        .Fields.Append .CreateField("LinkID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
    End With
    
    Set tbdTemp = daoDb.TableDefs("tmpEmailLinksRecipients")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxLinkID")
      With idxTemp
        .Fields.Append .CreateField("LinkID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      Set idxTemp = .CreateIndex("idxLinkRepID")
      With idxTemp
        .Fields.Append .CreateField("LinkID")
        .Fields.Append .CreateField("RecipientID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
  
  
    'MH20090520
    Set tbdTemp = daoDb.TableDefs("tmpLinkContent")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxContentIDSequence")
      With idxTemp
        .Fields.Append .CreateField("ContentID")
        .Fields.Append .CreateField("Sequence")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
    
    
    'Set tbdTemp = daoDb.TableDefs("tmpEmailLinksAttachments")
    'With tbdTemp
    '  Set idxTemp = .CreateIndex("idxLinkID")
    '  With idxTemp
    '    .Fields.Append .CreateField("LinkID")
    '    .Unique = False
    '  End With
    '  .Indexes.Append idxTemp
    '
    '  .Indexes.Refresh
    'End With
  
  
    'MH20040301
    Set tbdTemp = daoDb.TableDefs("tmpOutlookLinks")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("TableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
  
  
    Set tbdTemp = daoDb.TableDefs("tmpOutlookLinksDestinations")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxLinkID")
      With idxTemp
        .Fields.Append .CreateField("LinkID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
  
  
    Set tbdTemp = daoDb.TableDefs("tmpOutlookLinksColumns")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxLinkSeqID")
      With idxTemp
        .Fields.Append .CreateField("LinkID")
        .Fields.Append .CreateField("Sequence")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
    
    
    Set tbdTemp = daoDb.TableDefs("tmpOutlookFolders")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("TableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      Set idxTemp = .CreateIndex("idxFolderID")
      With idxTemp
        .Fields.Append .CreateField("FolderID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
  
  
    
    
    ' Create the indices for the local Expressions table.
    Set tbdTemp = daoDb.TableDefs("tmpExpressions")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxExprID")
      With idxTemp
        .Fields.Append .CreateField("exprID")
        .Fields.Append .CreateField("deleted")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxExprName")
      With idxTemp
        .Fields.Append .CreateField("name")
        .Unique = False
      End With
      .Indexes.Append idxTemp
          
      Set idxTemp = .CreateIndex("idxParent")
      With idxTemp
        .Fields.Append .CreateField("parentComponentID")
        .Fields.Append .CreateField("exprID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
          
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Expression Components table.
    Set tbdTemp = daoDb.TableDefs("tmpComponents")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxCompID")
      With idxTemp
        .Fields.Append .CreateField("componentID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxExprID")
      With idxTemp
        .Fields.Append .CreateField("exprID")
        .Fields.Append .CreateField("componentID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxCalcID")
      With idxTemp
        .Fields.Append .CreateField("calculationID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Views table.
    Set tbdTemp = daoDb.TableDefs("tmpViews")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxViewID")
      With idxTemp
        .Fields.Append .CreateField("ViewID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxViewName")
      With idxTemp
        .Fields.Append .CreateField("ViewName")
        .Fields.Append .CreateField("deleted")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxViewTableID")
      With idxTemp
        .Fields.Append .CreateField("ViewTableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local View columns table.
    Set tbdTemp = daoDb.TableDefs("tmpViewColumns")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxViewColID")
      With idxTemp
        .Fields.Append .CreateField("ViewID")
        .Fields.Append .CreateField("ColumnID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      Set idxTemp = .CreateIndex("idxViewID")
      With idxTemp
        .Fields.Append .CreateField("ViewID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
  
    ' Create the indices for the local View screens table.
    Set tbdTemp = daoDb.TableDefs("tmpViewScreens")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxViewID")
      With idxTemp
        .Fields.Append .CreateField("ViewID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxScreenID")
      With idxTemp
        .Fields.Append .CreateField("ScreenID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      Set idxTemp = .CreateIndex("idxViewScreen")
      With idxTemp
        .Fields.Append .CreateField("ViewID")
        .Fields.Append .CreateField("ScreenID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Module Setup table.
    Set tbdTemp = daoDb.TableDefs("tmpModuleSetup")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxModuleParameter")
      With idxTemp
        .Fields.Append .CreateField("ModuleKey")
        .Fields.Append .CreateField("ParameterKey")
        .Unique = False
'        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Module Related Columns table.
    Set tbdTemp = daoDb.TableDefs("tmpModuleRelatedColumns")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxModuleParameter")
      With idxTemp
        .Fields.Append .CreateField("ModuleKey")
        .Fields.Append .CreateField("ParameterKey")
        
        'JPD 20040121 Fault 7932
        '.Unique = true
        .Unique = False
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With
    
    'TM20020211 Fault 3487
    ' Create the indices for the local Mail Merge table.
    Set tbdTemp = daoDb.TableDefs("tmpMailMerge")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxMailMergeID")
      With idxTemp
        .Fields.Append .CreateField("MailMergeID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Workflows table.
    Set tbdTemp = daoDb.TableDefs("tmpWorkflows")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxWorkflowID")
      With idxTemp
        .Fields.Append .CreateField("ID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
  
      Set idxTemp = .CreateIndex("idxName")
      With idxTemp
        .Fields.Append .CreateField("name")
        .Fields.Append .CreateField("deleted")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
  
    ' Create the indices for the local Workflow Elements table.
    Set tbdTemp = daoDb.TableDefs("tmpWorkflowElements")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxElementID")
      With idxTemp
        .Fields.Append .CreateField("ID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
  
      Set idxTemp = .CreateIndex("idxWorkflowID")
      With idxTemp
        .Fields.Append .CreateField("workflowID")
        .Fields.Append .CreateField("ID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
  
    ' Create the indices for the local Workflow Links table.
    Set tbdTemp = daoDb.TableDefs("tmpWorkflowLinks")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxLinkID")
      With idxTemp
        .Fields.Append .CreateField("ID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
  
      Set idxTemp = .CreateIndex("idxWorkflowID")
      With idxTemp
        .Fields.Append .CreateField("workflowID")
        .Fields.Append .CreateField("ID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
  
    ' Create the indices for the local Workflow Element Items table.
    Set tbdTemp = daoDb.TableDefs("tmpWorkflowElementItems")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxItemID")
      With idxTemp
        .Fields.Append .CreateField("ID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
  
      Set idxTemp = .CreateIndex("idxElementID")
      With idxTemp
        .Fields.Append .CreateField("ElementID")
        .Fields.Append .CreateField("ID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
    
    ' Create the indices for the local Workflow Element Item Values table.
    Set tbdTemp = daoDb.TableDefs("tmpWorkflowElementItemValues")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxItemID")
      With idxTemp
        .Fields.Append .CreateField("itemID")
        .Fields.Append .CreateField("sequence")
        .Unique = True
      End With
      .Indexes.Append idxTemp
      
      .Indexes.Refresh
    End With

    ' Create the indices for the local Workflow Element Columns table.
    Set tbdTemp = daoDb.TableDefs("tmpWorkflowElementColumns")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxID")
      With idxTemp
        .Fields.Append .CreateField("ID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
  
      Set idxTemp = .CreateIndex("idxElementID")
      With idxTemp
        .Fields.Append .CreateField("ElementID")
        .Fields.Append .CreateField("ID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
  
    ' Create the indices for the local Workflow Element Validations table.
    Set tbdTemp = daoDb.TableDefs("tmpWorkflowElementValidations")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxID")
      With idxTemp
        .Fields.Append .CreateField("ID")
        .Unique = True
      End With
      .Indexes.Append idxTemp
  
      Set idxTemp = .CreateIndex("idxElementID")
      With idxTemp
        .Fields.Append .CreateField("ElementID")
        .Fields.Append .CreateField("ID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With
  
    ' Create the indices for the local Workflow Triggered Links table.
    Set tbdTemp = daoDb.TableDefs("tmpWorkflowTriggeredLinks")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxTableID")
      With idxTemp
        .Fields.Append .CreateField("TableID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
    
      .Indexes.Refresh
    End With

    Set tbdTemp = daoDb.TableDefs("tmpWorkflowTriggeredLinkColumns")
    With tbdTemp
      Set idxTemp = .CreateIndex("idxLinkID")
      With idxTemp
        .Fields.Append .CreateField("LinkID")
        .Unique = False
      End With
      .Indexes.Append idxTemp
  
      .Indexes.Refresh
    End With

    Set tbdTemp = Nothing
    Set idxTemp = Nothing
  
  End If

  ' Open temporary table recordsets.
  Set recTabEdit = daoDb.OpenRecordset("tmpTables", dbOpenTable)
  Set recSummaryEdit = daoDb.OpenRecordset("tmpSummary", dbOpenTable)
  Set recTableValidationEdit = daoDb.OpenRecordset("tmpTableValidations", dbOpenTable)
  Set recColEdit = daoDb.OpenRecordset("tmpColumns", dbOpenTable)
  Set recDiaryEdit = daoDb.OpenRecordset("tmpDiary", dbOpenTable)
  Set recContValEdit = daoDb.OpenRecordset("tmpControlValues", dbOpenTable)
  Set recRelEdit = daoDb.OpenRecordset("tmpRelations", dbOpenTable)
  Set recHistScrEdit = daoDb.OpenRecordset("tmpHistoryScreens", dbOpenTable)
  Set recScrEdit = daoDb.OpenRecordset("tmpScreens", dbOpenTable)
  Set recPageCaptEdit = daoDb.OpenRecordset("tmpPageCaptions", dbOpenTable)
  Set recCtrlEdit = daoDb.OpenRecordset("tmpControls", dbOpenTable)
  Set recPictEdit = daoDb.OpenRecordset("tmpPictures", dbOpenTable)
  Set recOrdEdit = daoDb.OpenRecordset("tmpOrders", dbOpenTable)
  Set recOrdItemEdit = daoDb.OpenRecordset("tmpOrderItems", dbOpenTable)
  
  'MH20000727
  Set recEmailAddrEdit = daoDb.OpenRecordset("tmpEmailAddresses", dbOpenTable)
  Set recEmailLinksEdit = daoDb.OpenRecordset("tmpEmailLinks", dbOpenTable)
  'Set recEmailAttachmentsEdit = daoDb.OpenRecordset("tmpEmailLinksAttachments", dbOpenTable)
  Set recEmailLinksColumnsEdit = daoDb.OpenRecordset("tmpEmailLinksColumns", dbOpenTable)
  Set recEmailRecipientsEdit = daoDb.OpenRecordset("tmpEmailLinksRecipients", dbOpenTable)
  
  'MH20090520
  Set recLinkContentEdit = daoDb.OpenRecordset("tmpLinkContent", dbOpenTable)
  
  'MH20040301
  Set recOutlookFolders = daoDb.OpenRecordset("tmpOutlookFolders", dbOpenTable)
  Set recOutlookLinks = daoDb.OpenRecordset("tmpOutlookLinks", dbOpenTable)
  Set recOutlookLinksColumns = daoDb.OpenRecordset("tmpOutlookLinksColumns", dbOpenTable)
  Set recOutlookLinksDestinations = daoDb.OpenRecordset("tmpOutlookLinksDestinations", dbOpenTable)


  Set recExprEdit = daoDb.OpenRecordset("tmpExpressions", dbOpenTable)
  Set recCompEdit = daoDb.OpenRecordset("tmpComponents", dbOpenTable)
  Set recViewEdit = daoDb.OpenRecordset("tmpViews", dbOpenTable)
  Set recViewColEdit = daoDb.OpenRecordset("tmpViewColumns", dbOpenTable)
  Set recViewScreens = daoDb.OpenRecordset("tmpViewScreens", dbOpenTable)
  Set recModuleSetup = daoDb.OpenRecordset("tmpModuleSetup", dbOpenTable)
  Set recModuleRelatedColumns = daoDb.OpenRecordset("tmpModuleRelatedColumns", dbOpenTable)
   
  'TM20020211 Fault 3487
  Set recMailMerge = daoDb.OpenRecordset("tmpMailMerge", dbOpenTable)
  
  Set recWorkflowEdit = daoDb.OpenRecordset("tmpWorkflows", dbOpenTable)
  Set recWorkflowElementEdit = daoDb.OpenRecordset("tmpWorkflowElements", dbOpenTable)
  Set recWorkflowLinkEdit = daoDb.OpenRecordset("tmpWorkflowLinks", dbOpenTable)
  Set recWorkflowElementItemEdit = daoDb.OpenRecordset("tmpWorkflowElementItems", dbOpenTable)
  Set recWorkflowElementItemValuesEdit = daoDb.OpenRecordset("tmpWorkflowElementItemValues", dbOpenTable)
  Set recWorkflowElementColumnEdit = daoDb.OpenRecordset("tmpWorkflowElementColumns", dbOpenTable)
  Set recWorkflowElementValidationEdit = daoDb.OpenRecordset("tmpWorkflowElementValidations", dbOpenTable)
  Set recWorkflowTriggeredLinks = daoDb.OpenRecordset("tmpWorkflowTriggeredLinks", dbOpenTable)
  Set recWorkflowTriggeredLinkColumns = daoDb.OpenRecordset("tmpWorkflowTriggeredLinkColumns", dbOpenTable)

  'Populate rdo tables collection
  'rdoCon.rdoTables.Refresh

  CreateTempTables = True
  Exit Function

ErrorTrap:
  CreateTempTables = False
  gobjProgress.Visible = False
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Name
  Err = False
  
End Function
Public Function DropTempTables() As Boolean
  On Error GoTo ErrorTrap
  
  ' Drop the local tables.
  recTabEdit.Close
  daoDb.Execute "DROP TABLE tmpTables"
  
  recTableValidationEdit.Close
  daoDb.Execute "DROP TABLE tmpTableValidations"
  
  recSummaryEdit.Close
  daoDb.Execute "DROP TABLE tmpSummary"
  
  recColEdit.Close
  daoDb.Execute "DROP TABLE tmpColumns"
  
  recDiaryEdit.Close
  daoDb.Execute "DROP TABLE tmpDiary"
  
  recContValEdit.Close
  daoDb.Execute "DROP TABLE tmpControlValues"
  
  recRelEdit.Close
  daoDb.Execute "DROP TABLE tmpRelations"
  
  recHistScrEdit.Close
  daoDb.Execute "DROP TABLE tmpHistoryScreens"
  
  recScrEdit.Close
  daoDb.Execute "DROP TABLE tmpScreens"
  
  recPageCaptEdit.Close
  daoDb.Execute "DROP TABLE tmpPageCaptions"
  
  recCtrlEdit.Close
  daoDb.Execute "DROP TABLE tmpControls"
  
  recPictEdit.Close
  daoDb.Execute "DROP TABLE tmpPictures"
  
  recOrdEdit.Close
  daoDb.Execute "DROP TABLE tmpOrders"
  
  recOrdItemEdit.Close
  daoDb.Execute "DROP TABLE tmpOrderItems"
   
  'MH20000727 Added Email Tables
  recEmailAddrEdit.Close
  daoDb.Execute "DROP TABLE tmpEmailAddresses"
  
  recEmailLinksEdit.Close
  daoDb.Execute "DROP TABLE tmpEmailLinks"
   
  recEmailLinksColumnsEdit.Close
  daoDb.Execute "DROP TABLE tmpEmailLinksColumns"
  
  recEmailRecipientsEdit.Close
  daoDb.Execute "DROP TABLE tmpEmailLinksRecipients"
 
  'MH20090520
  recLinkContentEdit.Close
  daoDb.Execute "DROP TABLE tmpLinkContent"

  'MH20040301
  recOutlookFolders.Close
  daoDb.Execute "DROP TABLE tmpOutlookFolders"

  recOutlookLinks.Close
  daoDb.Execute "DROP TABLE tmpOutlookLinks"

  recOutlookLinksColumns.Close
  daoDb.Execute "DROP TABLE tmpOutlookLinksColumns"

  recOutlookLinksDestinations.Close
  daoDb.Execute "DROP TABLE tmpOutlookLinksDestinations"
   
  recExprEdit.Close
  daoDb.Execute "DROP TABLE tmpExpressions"
  
  recCompEdit.Close
  daoDb.Execute "DROP TABLE tmpComponents"
  
  recViewEdit.Close
  daoDb.Execute "DROP TABLE tmpViews"
  
  recViewColEdit.Close
  daoDb.Execute "DROP TABLE tmpViewColumns"
  
  recViewScreens.Close
  daoDb.Execute "DROP TABLE tmpViewScreens"

  recModuleSetup.Close
  daoDb.Execute "DROP TABLE tmpModuleSetup"

  recModuleRelatedColumns.Close
  daoDb.Execute "DROP TABLE tmpModuleRelatedColumns"

  daoDb.Execute "DROP TABLE tmpSSIntranetLinks"
  daoDb.Execute "DROP TABLE tmpSSIHiddenGroups"
  daoDb.Execute "DROP TABLE tmpSSIViews"

  ' Payroll Integration Tables
  daoDb.Execute "DROP TABLE tmpAccordTransferFieldDefinitions"
  daoDb.Execute "DROP TABLE tmpAccordTransferFieldMappings"
  daoDb.Execute "DROP TABLE tmpAccordTransferTypes"

'  ' Fusion Tables
'  daoDb.Execute "DROP TABLE tmpFusionFieldDefinitions"
'  daoDb.Execute "DROP TABLE tmpFusionFieldMappings"
'  daoDb.Execute "DROP TABLE tmpFusionTypes"

  ' Mobile Navigation Tables
  daoDb.Execute "DROP TABLE tmpGroups"
  daoDb.Execute "DROP TABLE tmpmobileformlayout"
  daoDb.Execute "DROP TABLE tmpmobileformelements"
  daoDb.Execute "DROP TABLE tmpmobilegroupworkflows"

  recMailMerge.Close
  daoDb.Execute "DROP TABLE tmpMailMerge"
  
  recWorkflowEdit.Close
  daoDb.Execute "DROP TABLE tmpWorkflows"
  recWorkflowElementEdit.Close
  daoDb.Execute "DROP TABLE tmpWorkflowElements"
  recWorkflowLinkEdit.Close
  daoDb.Execute "DROP TABLE tmpWorkflowLinks"
  recWorkflowElementItemEdit.Close
  daoDb.Execute "DROP TABLE tmpWorkflowElementItems"
  recWorkflowElementItemValuesEdit.Close
  daoDb.Execute "DROP TABLE tmpWorkflowElementItemValues"
  recWorkflowElementColumnEdit.Close
  daoDb.Execute "DROP TABLE tmpWorkflowElementColumns"
  recWorkflowElementValidationEdit.Close
  daoDb.Execute "DROP TABLE tmpWorkflowElementValidations"
  recWorkflowTriggeredLinks.Close
  daoDb.Execute "DROP TABLE tmpWorkflowTriggeredLinks"
  recWorkflowTriggeredLinkColumns.Close
  daoDb.Execute "DROP TABLE tmpWorkflowTriggeredLinkColumns"
  
  DropTempTables = True
  Exit Function

ErrorTrap:
  DropTempTables = False
  gobjProgress.Visible = False
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Name
  Err = False
  
End Function

Public Function Login() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim psUser As String
  Dim psPassword As String
  
  If Application.LoggedIn Then
    Logout
  End If
  
  If Not Application.LoggedIn Then
    
    Load frmLogin
    Application.LoggedIn = frmLogin.OK

    If Not Application.LoggedIn Then
      frmLogin.Show vbModal
      Application.LoggedIn = frmLogin.OK
      'MH20061025 Fault 11625
      'Don't overwrite these as causes problem if you have forced changed password
      'gsUserName = frmLogin.txtUID.Text
      'gsPassword = frmLogin.txtPWD.Text
      If Not Application.LoggedIn Then
        Logout
      End If
    End If

    UnLoad frmLogin
    Set frmLogin = Nothing

  End If
  
  Login = Application.LoggedIn
  If Login Then
    Call AuditAccess("Log In", "System")
  End If
  Exit Function
  
ErrorTrap:
  Login = False
  Err = False
  
End Function

Public Function Logout() As Boolean
  On Error GoTo ErrorTrap

  If Not gADOCon Is Nothing Then
    If Application.AccessMode = accFull Then
      UnlockDatabase lckReadWrite
    End If
  End If


  If Application.LoggedIn Then
    
    Call AuditAccess("Log Out", "System")
    gADOCon.Close
    
    Application.LoggedIn = False
  End If
  
  Logout = (Not Application.LoggedIn)
 
  If Not (daoDb Is Nothing) Then
    daoDb.Close
    Set daoDb = Nothing
  End If
  
  ' Kills the ado connection
  If Not gADOCon Is Nothing Then
    If gADOCon.State = adStateOpen Then
      gADOCon.Close
    End If
  End If
  
  Set gADOCon = Nothing
  Set gobjHRProEngine = Nothing
  
  'Make sure temp ASR database does not already exist
  If Dir(Environ("TEMP") & "\" & gsTempDatabaseName) <> vbNullString Then
    Kill Environ("TEMP") & "\" & gsTempDatabaseName
  End If

  Exit Function
  
ErrorTrap:
  
  If Err.Number = 53 Then Resume Next
  MsgBox "Error whilst calling Application.Logout." & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, Application.Name
  Logout = False
  Err = False
  
End Function

Public Sub ActivateModules()
  ' Only show licenced modules in the menus.
  Dim sSQL As String
  Dim rsWorkflows As DAO.Recordset
  Dim sURL As String
  Dim sNewQueryString As String
  Dim sNewString As String
  Dim sOldString As String
  Dim sTemp As String
  Dim asWorkflows() As String
  Dim frmChangedPlatform As frmChangedPlatform
  Dim iLoop As Integer
  
  ReDim asWorkflows(2, 0)
  ' Column 0 = Workflow Name
  ' Column 1 = Original URL + QueryString
  ' Column 2 = New URL + QueryString
    
  glngEmailMethod = GetSystemSetting("Email", "Method", 1)
  gstrEmailProfile = GetSystemSetting("Email", "Profile", "")
  gstrEmailServer = GetSystemSetting("Email", "Server", "")
  gstrEmailAccount = GetSystemSetting("Email", "Account", "")

  gstrEmailAttachmentPath = GetSystemSetting("Email", "Attachment Path", "")
  glngEmailDateFormat = GetSystemSetting("Email", "Date Format", 103)
  gstrEmailTestAddr = GetSystemSetting("Email", "Test Messages", "hrpro@hrpro.co.uk")
'    gstrEmailEventLogToAddr = GetSystemSetting("Email", "Event Log Send", "SupportAtYourCompany@hrpro.co.uk")

  ' Desktop settings
  glngDesktopBitmapID = GetSystemSetting("DesktopSetting", "BitmapID", 0)
  glngDesktopBitmapLocation = GetSystemSetting("DesktopSetting", "BitmapLocation", 0)
  glngDeskTopColour = GetSystemSetting("DesktopSetting", "BackgroundColour", &H8000000C)

  ' Advanced Database settings
  gbManualRecursionLevel = GetSystemSetting("AdvancedDatabaseSetting", "ManualRecursion", False)
  giManualRecursionLevel = GetSystemSetting("AdvancedDatabaseSetting", "ManualRecursionLevel", giDefaultRecursionLevel)

  gbDisableSpecialFunctionAutoUpdate = GetSystemSetting("AdvancedDatabaseSetting", "SpecialFunctionAutoUpdate", False)
  gbReorganiseIndexesInOvernightJob = GetSystemSetting("AdvancedDatabaseSetting", "UpdateIndexesOvernight", True)
  
  '26/07/2001 MH
  '' CMG settings
  gbCMGExportUseCSV = GetSystemSetting("CMGExport", "UseCSV", False)
  gbCMGIgnoreBlanks = GetSystemSetting("CMGExport", "IgnoreBlanks", False)
  gbCMGReverseDateChanged = GetSystemSetting("CMGExport", "ReverseOutput", False)
  gbCMGExportFileCode = GetSystemSetting("CMGExport", "FileCode", True)
  gbCMGExportFieldCode = GetSystemSetting("CMGExport", "FieldCode", True)
  gbCMGExportLastChangeDate = GetSystemSetting("CMGExport", "LastChange", True)
  giCMGExportFileCodeSize = GetSystemSetting("CMGExport", "FileCodeSize", 6)
  giCMGEXportRecordIDSize = GetSystemSetting("CMGExport", "RecordIdentifierSize", 11)
  giCMGExportFieldCodeSize = GetSystemSetting("CMGExport", "FieldCodeSize", 10)
  giCMGExportOutputColumnSize = GetSystemSetting("CMGExport", "OutputColumnSize", 53)
  giCMGExportLastChangeDateSize = GetSystemSetting("CMGExport", "LastChangeSize", 8)
  'NPG20090313 Fault 13595
  giCMGExportFileCodeOrderID = GetSystemSetting("CMGExport", "FileCodeOrderID", 0)
  giCMGEXportRecordIDOrderID = GetSystemSetting("CMGExport", "RecordIdentifierOrderID", 0)
  giCMGExportFieldCodeOrderID = GetSystemSetting("CMGExport", "FieldCodeOrderID", 0)
  giCMGExportOutputColumnOrderID = GetSystemSetting("CMGExport", "OutputColumnOrderID", 0)
  giCMGExportLastChangeDateOrderID = GetSystemSetting("CMGExport", "LastChangeOrderID", 0)

  glngExpressionViewColours = GetSystemSetting("ExpressionBuilder", "ViewColours", EXPRESSIONBUILDER_COLOUROFF)
  glngExpressionViewNodes = GetSystemSetting("ExpressionBuilder", "NodeSize", EXPRESSIONBUILDER_NODESMINIMIZE)

  gbMaximizeScreens = GetSystemSetting("General", "MaximizeScreens", False)
  
  ' Which columns are to be shown
  LoadShowWhichColumns

  ' SQL 2005 Process Info
  glngProcessMethod = GetSystemSetting("ProcessAccount", "Mode", 1)
  
  'load the overnight job schedule variables
  glngOvernightJobTime = GetSystemSetting("overnight", "time", 30000)

  'MH20040301
  glngAMStartTime = val(Replace(GetSystemSetting("Outlook", "AMStartTime", 900), ":", ""))
  glngAMEndTime = val(Replace(GetSystemSetting("Outlook", "AMEndTime", 1230), ":", ""))
  glngPMStartTime = val(Replace(GetSystemSetting("Outlook", "PMStartTime", 1330), ":", ""))
  glngPMEndTime = val(Replace(GetSystemSetting("Outlook", "PMEndTime", 1700), ":", ""))
  
  ' Postcode Integration modules
  gbAFDEnabled = IsModuleEnabled(modAFD)
  gbQAddressEnabled = IsModuleEnabled(modQAddress)
   
  If IsModuleEnabled(modWorkflow) And (gfDatabaseServerChanged Or gfWFCredentialsChanged) Then
    sURL = GetWorkflowURL
    
    If Len(sURL) > 0 Then
      sSQL = "SELECT tmpWorkflows.name," & _
        "   tmpWorkflows.ID," & _
        "   tmpWorkflows.queryString" & _
        " FROM tmpWorkflows " & _
        " WHERE tmpWorkflows.initiationType = " & CStr(WORKFLOWINITIATIONTYPE_EXTERNAL)
  
      Set rsWorkflows = daoDb.OpenRecordset(sSQL)
      With rsWorkflows
        If Not (.EOF And .BOF) Then
          Do Until .EOF
            sOldString = IIf(Len(!queryString) > 0, sURL & "?" & !queryString, "")
            
            sNewQueryString = GetWorkflowQueryString(!ID * -1, -1)
  
            sSQL = "UPDATE tmpWorkflows" & _
              " SET changed = TRUE," & _
              "   queryString = '" & Replace(sNewQueryString, "'", "''") & "'" & _
              " WHERE ID = " & CStr(!ID)
            daoDb.Execute sSQL

            sNewString = IIf((Len(sNewQueryString) > 0), sURL & "?" & sNewQueryString, "")

            If (sOldString <> sNewString) _
              And (Len(sOldString) + Len(sNewString) > 0) Then
              ReDim Preserve asWorkflows(2, UBound(asWorkflows, 2) + 1)
              asWorkflows(0, UBound(asWorkflows, 2)) = !Name
              asWorkflows(1, UBound(asWorkflows, 2)) = sOldString
              asWorkflows(2, UBound(asWorkflows, 2)) = sNewString
            End If

            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsWorkflows = Nothing
    End If
    
    If UBound(asWorkflows, 2) > 0 Then
      Set frmChangedPlatform = New frmChangedPlatform
      frmChangedPlatform.ResetList

      For iLoop = 1 To UBound(asWorkflows, 2)
        frmChangedPlatform.AddToList asWorkflows(0, iLoop), _
          IIf(Len(asWorkflows(1, iLoop)) > 0, asWorkflows(1, iLoop), "<none>"), _
          IIf(Len(asWorkflows(2, iLoop)) > 0, asWorkflows(2, iLoop), "<none>")
      Next iLoop

      frmChangedPlatform.Width = (3 * Screen.Width / 4)
      frmChangedPlatform.Height = (Screen.Height / 2)

      Screen.MousePointer = vbDefault
      Call LoadSkin(frmChangedPlatform, frmSysMgr.SkinFramework1)
      frmChangedPlatform.ShowMessage 1
      
      UnLoad frmChangedPlatform
      Set frmChangedPlatform = Nothing
    End If
  End If

End Sub

Private Function ChangeConnectParam(strConString As String, strParameter As String, strNewValue As String) As String

  'MH20010704 Will subsitute a parameter in the connection string for a new value
  'e.g. pass in "APP=" and a new value and function will return new connection string

  Dim strParamArray As Variant
  Dim lngCount As Long
  
  strParamArray = Split(strConString, ";")
  

  For lngCount = 0 To UBound(strParamArray)
    
    If Left(strParamArray(lngCount), Len(strParameter)) = strParameter Then
      strParamArray(lngCount) = "APP=" & strNewValue
      Exit For
    End If
  
  Next

  ChangeConnectParam = Join(strParamArray, ";")

End Function

' Loads which columns are displayed in the listviews
Public Sub LoadShowWhichColumns()

  gbRememberDBColumnsView = GetSystemSetting("General", "RememberColumnsView", False)
  
  ' Load the Data manager column types
  Set gpropShowColumns_DataMgr = New SystemMgr.Properties
  gpropShowColumns_DataMgr.Add "Name", GetSystemSetting("ShowColumn_DataMan", "Name", True)
  gpropShowColumns_DataMgr.Add "Data Type", GetSystemSetting("ShowColumn_DataMan", "Data Type", True)
  gpropShowColumns_DataMgr.Add "Size", GetSystemSetting("ShowColumn_DataMan", "Size", True)
  gpropShowColumns_DataMgr.Add "Decimals", GetSystemSetting("ShowColumn_DataMan", "Decimals", True)
  gpropShowColumns_DataMgr.Add "Display", GetSystemSetting("ShowColumn_DataMan", "Display", True)
  gpropShowColumns_DataMgr.Add "Column Type", GetSystemSetting("ShowColumn_DataMan", "Column Type", True)
  gpropShowColumns_DataMgr.Add "Control Type", GetSystemSetting("ShowColumn_DataMan", "Control Type", True)
  gpropShowColumns_DataMgr.Add "Read Only", GetSystemSetting("ShowColumn_DataMan", "Read Only", True)
  gpropShowColumns_DataMgr.Add "Audit", GetSystemSetting("ShowColumn_DataMan", "Audit", True)
  gpropShowColumns_DataMgr.Add "Multi-line", GetSystemSetting("ShowColumn_DataMan", "Multi-line", False)
  gpropShowColumns_DataMgr.Add "Case", GetSystemSetting("ShowColumn_DataMan", "Case", False)
  gpropShowColumns_DataMgr.Add "Default Value", GetSystemSetting("ShowColumn_DataMan", "Default Value", False)
  gpropShowColumns_DataMgr.Add "Text Alignment", GetSystemSetting("ShowColumn_DataMan", "Text Alignment", False)
  gpropShowColumns_DataMgr.Add "Duplicate Check", GetSystemSetting("ShowColumn_DataMan", "Duplicate Check", False)
  gpropShowColumns_DataMgr.Add "Mandatory", GetSystemSetting("ShowColumn_DataMan", "Mandatory", True)
  gpropShowColumns_DataMgr.Add "Unique in Table", GetSystemSetting("ShowColumn_DataMan", "Unique in Table", False)
  gpropShowColumns_DataMgr.Add "Unique in Siblings", GetSystemSetting("ShowColumn_DataMan", "Unique in Siblings", False)
  gpropShowColumns_DataMgr.Add "Mask", GetSystemSetting("ShowColumn_DataMan", "Mask", False)
  gpropShowColumns_DataMgr.Add "Custom Validation", GetSystemSetting("ShowColumn_DataMan", "Custom Validation", False)
  gpropShowColumns_DataMgr.Add "Diary Links", GetSystemSetting("ShowColumn_DataMan", "Diary Links", True)
  'gpropShowColumns_DataMgr.Add "Email Links", GetSystemSetting("ShowColumn_DataMan", "Email Links", True)
  
  'NHRD23072003 Fault 6207
  If IsModuleEnabled(modAFD) Then
    gpropShowColumns_DataMgr.Add "AFD Postcode", GetSystemSetting("ShowColumn_DataMan", "AFD Postcode", False)
  End If
    
  If IsModuleEnabled(modQAddress) Then
    gpropShowColumns_DataMgr.Add "Quick Address", GetSystemSetting("ShowColumn_DataMan", "Quick Address", False)
  End If
  
  'NHRD29072003 Fault 6208 Added the ability to show and store Use1000Separator and Trimming properties
  gpropShowColumns_DataMgr.Add "Use 1000 Separator", GetSystemSetting("ShowColumn_DataMan", "Use 1000 Separator", True)
  gpropShowColumns_DataMgr.Add "Trimming", GetSystemSetting("ShowColumn_DataMan", "Trimming", True)
  
  ' Load the Data manager (table) column types
  Set gpropShowColumns_DataMgrTable = New SystemMgr.Properties
  gpropShowColumns_DataMgrTable.Add "Name", GetSystemSetting("ShowColumn_DataManTable", "Name", True)
  gpropShowColumns_DataMgrTable.Add "Type", GetSystemSetting("ShowColumn_DataManTable", "Type", True)
  gpropShowColumns_DataMgrTable.Add "Primary Order", GetSystemSetting("ShowColumn_DataManTable", "Primary Order", True)
  gpropShowColumns_DataMgrTable.Add "Record Description", GetSystemSetting("ShowColumn_DataManTable", "Record Description", True)
  gpropShowColumns_DataMgrTable.Add "Default Email", GetSystemSetting("ShowColumn_DataManTable", "Default Email", True)
  gpropShowColumns_DataMgrTable.Add "Email Links", GetSystemSetting("ShowColumn_DataManTable", "Email Links", True)
  gpropShowColumns_DataMgrTable.Add "Calendar Links", GetSystemSetting("ShowColumn_DataManTable", "Calendar Links", True)
    
  If IsModuleEnabled(modWorkflow) Then
    gpropShowColumns_DataMgrTable.Add "Workflow Links", GetSystemSetting("ShowColumn_DataManTable", "Workflow Links", True)
  End If
  
  ' Load the Picture manager column types
  Set gpropShowColumns_PictMgr = New SystemMgr.Properties
  gpropShowColumns_PictMgr.Add "Name", GetSystemSetting("ShowColumn_PictMan", "Name", True)
  gpropShowColumns_PictMgr.Add "Type", GetSystemSetting("ShowColumn_PictMan", "Type", True)
  gpropShowColumns_PictMgr.Add "Height", GetSystemSetting("ShowColumn_PictMan", "Height", True)
  gpropShowColumns_PictMgr.Add "Width", GetSystemSetting("ShowColumn_PictMan", "Width", True)

  ' Load the View manager column types
  Set gpropShowColumns_ViewMgr = New SystemMgr.Properties
  gpropShowColumns_ViewMgr.Add "Name", GetSystemSetting("ShowColumn_ViewMan", "Name", True)
  gpropShowColumns_ViewMgr.Add "Description", GetSystemSetting("ShowColumn_ViewMan", "Description", True)

End Sub

' Create the query defs on the temp Access DB for the new Phoenix engine to work.
Public Function CreateQueryDefs() As Boolean

  On Error GoTo ErrorTrap

  Dim sSQL As String
  Dim bOK As Boolean

  bOK = True

  ' spadmin_gettables
  sSQL = "SELECT t.tableid AS ID" & _
      ", t.tablename AS name" & _
      ", 1 AS type" & _
      ", '' AS description" & _
      ", 0 AS isremoteview" & _
      ", t.[tabletype]" & _
      ", t.[RecordDescExprID] AS recorddescriptionid" & _
      ", t.[AuditInsert] AS auditinsert" & _
      ", t.[AuditDelete] AS auditdelete" & _
      ", t.[DefaultEmailID] AS defaultemailid" & _
      ", t.[DefaultOrderID] AS defaultorderid" & _
      ", IIF(t.[deleted]=-1, 8 , IIF(t.[new]=-1,4 ,IIF(t.[changed]=-1,16, 2))) AS [state]" & _
      " FROM tmpTables t;"
  daoDb.CreateQueryDef "spadmin_gettables", sSQL

  ' spadmin_getcolumns
  sSQL = "SELECT c.columnid as [ID]" & _
      ", c.columnname as [name]" & _
      ", 2 as [type]" & _
      ", '' AS [description]" & _
      ", c.[calcExprID] AS [calcid]" & _
      ", c.[datatype] as [datatype]" & _
      ", c.[size], [decimals]" & _
      ", c.[audit], [mandatory], [multiline] " & _
      ", c.[dfltvalueexprid] as [defaultcalcid]" & _
      ", c.[defaultvalue] as [defaultvalue]" & _
      ", c.[convertcase] as [case]" & _
      ", c.[readOnly] as [isreadonly]" & _
      ", c.[uniquechecktype]" & _
      ", c.[Alignment], c.[trimming], c.[calculateifempty]" & _
      ", IIF(c.[deleted]=-1, 8 , IIF(c.[new]=-1,4 ,IIF(c.[changed]=-1,16, 2))) AS [state]" & _
      ", c.TableID" & _
      " FROM tmpColumns c WHERE c.columntype <> 3" & _
      " ORDER BY c.[columnname];"
  daoDb.CreateQueryDef "spadmin_getcolumns", sSQL

  ' spadmin_getrelations
  sSQL = "SELECT r.ParentID, r.ChildID, p.TableName AS [ParentName], c.TableName AS [ChildName]" & _
    " FROM tmpTables c" & _
    " INNER JOIN (tmpTables p" & _
    " INNER JOIN tmpRelations r ON p.tableid = r.ParentID) ON c.tableid = r.ChildID;"
  daoDb.CreateQueryDef "spadmin_getrelations", sSQL

  ' spadmin_getviews
  sSQL = "SELECT v.[viewid] AS [id], v.[viewname] AS [name], viewdescription AS [description], v.[viewtableid] AS [tableid], v.[expressionid] AS [filterid]" & _
    " FROM tmpviews v" & _
    " ORDER BY v.[viewname];"
  daoDb.CreateQueryDef "spadmin_getviews", sSQL

  'spadmin_getviewitems
  sSQL = "SELECT [viewid], [ColumnID]" & _
    " FROM tmpviewcolumns" & _
    " WHERE InView = true;"
  daoDb.CreateQueryDef "spadmin_getviewitems", sSQL

  ' spadmin_getexpressions
  sSQL = "SELECT [exprid] as [ID]" & _
    " , [Name]" & _
    " , [Type]" & _
    " , [Description] " & _
    " , [ReturnType] " & _
    " , [ReturnSize] AS [size]" & _
    " , [ReturnDecimals] as [decimals]" & _
    " , IIF([deleted]=-1, 8 , IIF([new]=-1,4 ,IIF([changed]=-1,16, 2))) AS [state]" & _
    " , [TableID]" & _
    " FROM tmpExpressions" & _
    " WHERE ParentComponentID = 0 AND type < 10;"
  daoDb.CreateQueryDef "spadmin_getexpressions", sSQL

  ' spadmin_getcomponent_base
  sSQL = "SELECT c.componentid AS [componentid], c.[type] AS [subtype], e.[Name] AS [name]" & _
      " , 1 AS [level], 1 AS [sequence]" & _
      " , e.ReturnType AS [returntype], e.ReturnSize AS [returnsize], e.ReturnDecimals AS [returndecimals]" & _
      " , IIF(ISNULL(c.FieldPassBy),1,c.FieldPassBy)-1 AS [iscolumnbyreference]" & _
      " , IIF(ISNULL(c.FunctionID),0,c.FunctionID) AS [functionid]" & _
      " , IIF(ISNULL(c.OperatorID),0,c.OperatorID) AS [operatorid]" & _
      " , IIF(ISNULL(c.fieldtableid),0,c.fieldtableid) AS [tableid]" & _
      " , IIF(ISNULL(c.FieldColumnID),0,c.FieldColumnID) AS [columnid]" & _
      " , IIF(ISNULL(c.FieldSelectionRecord),0,c.FieldSelectionRecord) AS [columnaggregiatetype]" & _
      " , IIF(ISNULL(c.FieldSelectionLine),0,c.FieldSelectionLine) AS [specificline]" & _
      " , IIF(ISNULL(c.FieldSelectionOrderID),0,c.FieldSelectionOrderID) AS [columnorderid]" & _
      " , IIF(ISNULL(c.[FieldSelectionFilter]),0,c.FieldSelectionFilter) AS [columnfilterid]" & _
      " , IIF(ISNULL(c.[CalculationID]),0,c.CalculationID) AS [calculationid]" & _
      " , IIF(ISNULL(c.[FilterID]),0,c.FilterID) AS [filterid]" & _
      " , IIF(ISNULL(c.[ValueType]),0,c.ValueType) AS [valuetype]" & _
      " , IIF(ISNULL(c.[ValueCharacter]),'',c.ValueCharacter) AS [valuestring]" & _
      " , IIF(ISNULL(c.[ValueNumeric]),0,c.ValueNumeric) AS [valuenumeric]" & _
      " , IIF(ISNULL(c.[ValueLogic]),0,c.ValueLogic) AS [valuelogic]" & _
      " , IIF(ISNULL(c.[ValueDate]),0,c.ValueDate) AS [valuedate]" & _
      " , IIF(ISNULL(c.[LookupTableID]),0,c.LookupTableID) AS [LookupTableID]" & _
      " , IIF(ISNULL(c.[LookupColumnID]),0,c.LookupColumnID) AS [LookupColumnID]" & _
      " , 0 AS [isevaluated] , c.exprid AS ExpressionID" & _
      " FROM tmpExpressions e" & _
      " INNER JOIN tmpComponents c ON c.exprID = e.exprID" & _
      " ORDER BY c.componentid ASC;"
  daoDb.CreateQueryDef "spadmin_getcomponent_base", sSQL

  ' spadmin_getcomponent_function
  sSQL = "SELECT e.ExprID AS componentid, 9 AS subtype, e.name, 1 AS [level], e.[exprid] AS sequence, e.returntype" & _
      " , e.ReturnSize AS returnsize, e.ReturnDecimals AS returndecimals, c.[FunctionID] AS functionid, 0 AS operatorid" & _
      " , 0 AS [tableid], 0 AS columnid, 0 AS iscolumnbyreference, 0 AS columnaggregiatetype, 0 AS columnorderid, 0 AS columnfilterid" & _
      " , 0 AS [specificline]" & _
      " , 0 AS calculationid, 0 AS filterid, 0 AS valuetype, '' AS valuestring, 0 AS valuenumeric, 0 AS valuelogic, c.[ValueDate], 0 AS lookuptableid, 0 AS lookupcolumnid" & _
      " , IIF(e.type=1 AND e.returnType=3,1,0) AS isevaluated, c.componentid AS ExpressionID" & _
      " FROM tmpComponents AS c" & _
      " INNER JOIN tmpExpressions AS e ON e.ParentComponentID = c.ComponentID" & _
      " ORDER BY e.[ExprID];"
  daoDb.CreateQueryDef "spadmin_getcomponent_function", sSQL

  ' spadmin_getorder
  sSQL = "SELECT OrderID, Name, TableID, Type FROM tmpOrders;"
  daoDb.CreateQueryDef "spadmin_getorders", sSQL

  ' spadmin_getorderitems
  sSQL = "SELECT [OrderID], [ColumnID], [Type], [Sequence], [Ascending] FROM tmpOrderItems;"
  daoDb.CreateQueryDef "spadmin_getorderitems", sSQL

  ' spadmin_getmodulesetup
  sSQL = "SELECT [ModuleKey], [ParameterKey], [ParameterValue] AS [value]" & _
      ",IIF([ParameterType] = ""PType_ColumnID"" , 2, IIF([ParameterType] = ""PType_TableID"", 1, IIF([ParameterType] = ""PType_ScreenID"" , 14, 0))) As SubType" & _
      " FROM tmpModuleSetup;"
  daoDb.CreateQueryDef "spadmin_getmodulesetup", sSQL

  ' spadmin_getvalidations
  sSQL = "SELECT 1 AS [ValidationType], * FROM tmpColumns WHERE [Duplicate] = -1" & _
         " UNION " & _
         " SELECT 2 AS [ValidationType], * FROM tmpColumns WHERE [UniqueCheckType] = -1" & _
         " UNION " & _
         " SELECT 3 AS [ValidationType], * FROM tmpColumns WHERE [ChildUniqueCheck] = -2" & _
         " UNION " & _
         " SELECT 4 AS [ValidationType], * FROM tmpColumns WHERE [Mandatory] = -1;"
  daoDb.CreateQueryDef "spadmin_getvalidations", sSQL

  ' spadmin_getdescriptions
  sSQL = "SELECT tmpExpressions.[exprid] AS ID, tmpExpressions.[Name], tmpExpressions.[Type], tmpExpressions.[Description], tmpExpressions.[ReturnType], tmpExpressions.[ReturnSize] AS [size], tmpExpressions.[ReturnDecimals] AS decimals, IIf([tmpExpressions.deleted]=-1,8,IIf([tmpExpressions.new]=-1,4,IIf([tmpExpressions.changed]=-1,16,2))) AS state, tmpExpressions.[TableID] " & _
         "FROM tmpExpressions " & _
         "INNER JOIN tmptables ON tmptables.RecordDescExprID = tmpExpressions.ExprID " & _
         "WHERE tmpExpressions.[ParentComponentID]=0;"
  daoDb.CreateQueryDef "spadmin_getdescriptions", sSQL

  ' spadmin_getmasks
  sSQL = "SELECT DISTINCT [exprid] AS ID, [Name], [Type], [Description], [ReturnType], [ReturnSize] AS [size], [ReturnDecimals] AS decimals, IIF([tmpexpressions.deleted]=-1, 8 , IIF([tmpexpressions.new]=-1,4 ,IIF([tmpexpressions.changed]=-1,16, 2))) AS state, tmpcolumns.TableID" & vbNewLine & _
         "FROM tmpcolumns" & vbNewLine & _
         "INNER JOIN tmpexpressions ON tmpcolumns.[lostfocusexprid] = tmpexpressions.[exprid];"
  daoDb.CreateQueryDef "spadmin_getmasks", sSQL

TidyUpAndExit:
  CreateQueryDefs = bOK
  Exit Function
  
ErrorTrap:
  bOK = False
  MsgBox "Error creating spadmin functions on the TempDB"
  Resume TidyUpAndExit

End Function

' Add an itemkey and data to a combobox
Public Function AddItemToComboBox(ByRef combo As ComboBox, ItemText As String, ItemData As Variant) As Integer

  Dim lngKey As Integer

  combo.AddItem ItemText
  lngKey = combo.NewIndex
  
  combo.ItemData(lngKey) = ItemData
 
  AddItemToComboBox = lngKey

End Function

' Add an item to a listbox
Public Function AddItemToListbox(ByRef TheListbox As ListBox, ItemText As String, ItemData As Variant, IsSelected As Boolean) As Integer

  Dim lngKey As Integer

  TheListbox.AddItem ItemText
  lngKey = TheListbox.NewIndex
          
  TheListbox.ItemData(lngKey) = ItemData
  TheListbox.Selected(lngKey) = IsSelected

  AddItemToListbox = lngKey

End Function

' Disable/enable all open forms
Public Sub EnableOpenForms(ByVal pbEnabled As Boolean)
 
 Dim lngCount As Long
 
  ' Disable control while we save changes
  EnableCloseButton frmSysMgr.hWnd, pbEnabled
  For lngCount = 0 To (Forms.Count - 1)
      Forms(lngCount).Enabled = pbEnabled
  Next lngCount

End Sub

Public Sub EditMobileDesigner()

  On Error GoTo ErrorTrap
  
  Screen.MousePointer = vbHourglass
  
  Dim service As New MobileDesignerSerivce
  service.InitialiseForVB6 (Environ("TEMP") & "\" & gsTempDatabaseName)
  Set service = Nothing
    
  Dim changesMade As Boolean
  Dim frm As New DesignerForm
  frm.ReadOnly = (Application.AccessMode <> accFull And Application.AccessMode <> accSupportMode)
  
  Screen.MousePointer = vbDefault
  changesMade = frm.ShowForVB6
      
  If changesMade Then
    Application.Changed = True
    frmSysMgr.RefreshMenu True
  End If
  
  Exit Sub
ErrorTrap:
  MsgBox "An error occurred while showing the mobile designer." & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, Application.Name
End Sub

Public Sub AttemptReLogin()

  Dim bOK As Boolean
  Dim iRetryCount As Integer
  Dim iAnswer As Integer
 
  iRetryCount = 0
  bOK = False

  Do While Not bOK
    iRetryCount = iRetryCount + 1
    iAnswer = MsgBox("Your network connectivity has been lost." & vbCrLf & vbCrLf _
      & "Would you like to attempt to automatically relogin? " & vbCrLf & vbCrLf _
      & "If this happens on a regular basis please contact your system administrator as it is likely that there are some underlying network issues." _
      & " Attempt #" & iRetryCount _
    , vbCritical + vbYesNo, App.Title)
    
    If iAnswer = 6 Then
      bOK = AttemptConnection
    Else
      GoTo Quit
    End If
  Loop
  
  Exit Sub

Quit:

  Set gADOCon = Nothing
  Application.LoggedIn = False
  Application.Changed = False
  UnLoad frmSysMgr

End Sub

Public Function AttemptConnection() As Boolean

  On Error GoTo ErrorTrap:

  Screen.MousePointer = vbHourglass

  If gADOCon.State = adStateOpen Then
    gADOCon.Close
  End If
  gADOCon.Open
  
  Screen.MousePointer = vbDefault
  AttemptConnection = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  AttemptConnection = False
  GoTo TidyUpAndExit

End Function
