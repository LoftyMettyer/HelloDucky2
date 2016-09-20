Attribute VB_Name = "modSave"
Option Explicit

Private mfrmUse As frmUsage

Function SaveChanges(Optional pfRefreshDatabase As Boolean) As Boolean
  On Error GoTo ErrorTrap
 
  Dim fOK As Boolean
  Dim fInTransaction As Boolean
  Dim sErrMsg As String
  Dim alngExpressions() As Long
  Dim bFailed As Boolean
  Dim bLockedOK As Boolean
  
  fInTransaction = False
  bFailed = False
  fOK = True
  
  frmSysMgr.tmrKeepAlive.Enabled = False
  
  ReDim alngExpressions(2, 0)

  OutputCurrentProcess "Start of save process", True

  If IsMissing(pfRefreshDatabase) Then
    pfRefreshDatabase = False
  End If
  pfRefreshDatabase = (pfRefreshDatabase Or gfRefreshStoredProcedures)
   
  ' Disable control while we save changes
  EnableOpenForms False
  
  If fOK Then
    
    OutputCurrentProcess "Refreshing Database = " & IIf(pfRefreshDatabase, "'True'", "'False'")
    OutputCurrentProcess "Changed Table Name = " & IIf(Application.ChangedTableName, "'True'", "'False'")
    OutputCurrentProcess "Changed View Name = " & IIf(Application.ChangedViewName, "'True'", "'False'")
    OutputCurrentProcess "Changed Column Name = " & IIf(Application.ChangedColumnName, "'True'", "'False'")
    OutputCurrentProcess "Changed Overnight Job Schedule = " & IIf(Application.ChangedOvernightJobSchedule, "'True'", "'False'")

    OutputCurrentProcess vbNullString
    OutputCurrentProcess "Display Progress Bar"
    ' Display the progress bar.
    With gobjProgress
      .AVI = dbSave
      .Caption = Application.Name
      .MainCaption = "Saving Changes"
      .NumberOfBars = 2
      .Bar1Value = 0
      .Bar1MaxValue = 32
      .Bar2Value = 0
      .Bar1Caption = "Updating the server database..."
      .Time = False
      .Cancel = True
      .OpenProgress
    End With
    Screen.MousePointer = vbHourglass
  
    If Forms.Count > 0 Then
      frmSysMgr.StatusBar1.SimpleText = vbNullString
      frmSysMgr.StatusBar1.Refresh
    End If
    
  End If
    
  'JDM - 19/02/02 - Fault 3504 - Moved the quick checks before calling the support mode, or checking for other users
  If fOK Then
    DoEvents
    OutputCurrentProcess "Initialising Database Scripting"
    OutputCurrentProcess2 "Orders", 4
       
    fOK = QuickChecks_1
    gobjProgress.UpdateProgress2
    fOK = fOK And Not gobjProgress.Cancelled

    ' Quick checks - part 2. The ones that handle their own messages.
    If fOK Then
      OutputCurrentProcess2 "Descriptions", 4
      fOK = QuickChecks_2
      gobjProgress.UpdateProgress2
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' Quick checks - part 3. Check if the child tables have parents.
    If fOK Then
      OutputCurrentProcess2 "Orphan Check", 4
      fOK = QuickChecks_3
      gobjProgress.UpdateProgress2
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' Quick checks - part 4. Make sure a sysprocesses account has been defined for SQL 2005
    If fOK Then
      OutputCurrentProcess2 "Processing Check", 4
      RegenerateProcessAccount
      fOK = QuickChecks_4
      gobjProgress.UpdateProgress2
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
  End If
 
  ' Only allow save to proceed if they have the correct support code
  If fOK And Application.AccessMode = accSupportMode Then
    gobjProgress.Visible = False
    frmSupportMode.Show vbModal
    fOK = Not frmSupportMode.Cancelled
    UnLoad frmSupportMode
    Set frmSupportMode = Nothing
    gobjProgress.Visible = fOK
  End If
  
  ' Check there are no other users in the system
  If fOK Then
    fOK = SaveChanges_LogoutCheck(True)
    If fOK Then
      LockDatabase (lckSaving)
    Else
      MsgBox "Save process cancelled.", vbOKOnly + vbExclamation, Application.Name
    End If
    gobjProgress.Visible = fOK
  End If
  
  If fOK Then
    bLockedOK = True
  End If
  
  If fOK Then
    gobjProgress.ResetBar2
    OutputCurrentProcess "Reading Permissions"
    gobjProgress.UpdateProgress False
    DoEvents
    fOK = ReadPermissions(sErrMsg)
    fOK = fOK And Not gobjProgress.Cancelled
  End If
  
  ' Set database compatability (must be done outside of a transaction)
  If fOK Then
    gobjProgress.ResetBar2
    OutputCurrentProcess "Setting Database Compatibility"
    gobjProgress.UpdateProgress False
    DoEvents
    fOK = SetDatabaseCompatability
    fOK = fOK And Not gobjProgress.Cancelled
  End If
  
  If fOK Then
    OutputCurrentProcess "Begining Transaction"
  '  OutputCurrentProcess vbnullstring
  
    ' Begin transactions of data from local to remote databases.
    gADOCon.BeginTrans
    fInTransaction = True
                 
    ' Apply any post save hotfixes
    If fOK Then
      fOK = ApplyHotfixes(BEFORESAVE)
    End If
         
    ' Tidy up existing temporary tables/procedures/udfs
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Initialising .NET System Framework"
      gobjHRProEngine.Initialise
      gobjHRProEngine.Options.VersionUpgraded = gfRefreshStoredProcedures
      gobjProgress.UpdateProgress False
      DoEvents
      
      ' Cleanup
      fOK = CleanupDatabase

      ' Schema Binding Prep
      fOK = fOK And DropViews
      fOK = fOK And DropHierarchySpecifics
      
      fOK = fOK And Not gobjProgress.Cancelled
    End If
      
      
    ' Initialise the .NET engine
    If fOK Then
      Set gobjHRProEngine.CommitDB = gADOCon
      Set gobjHRProEngine.MetadataDB = daoDb
      fOK = gobjHRProEngine.PopulateObjects
      gobjHRProEngine.Options.RefreshObjects = True
    End If
    
    
    'MH20010227 Messageboxes are now produced within 'SaveTables' etc.
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Generating Tables"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveTables(pfRefreshDatabase, mfrmUse)
      fOK = fOK And Not gobjProgress.Cancelled
    End If


    ' Generate indexes
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Generate Indexes"
      gobjProgress.UpdateProgress False
      DoEvents

      OutputCurrentProcess2 "Primary keys", 2
      fOK = CreatePrimaryKeysForTables
      gobjProgress.UpdateProgress2

      OutputCurrentProcess2 "Foreign keys"
      fOK = fOK And CreateChildTableForeignKeys
      fOK = fOK And Not gobjProgress.Cancelled

      If fOK Then
        fOK = CreateHierarchyIndexes
      End If

    End If
      
      
    'Create Payroll lookup column calculations
    If IsModuleEnabled(modAccord) Then
      If fOK Then
        DoEvents
        OutputCurrentProcess "Payroll Integration"
        gobjProgress.UpdateProgress False
        OutputCurrentProcess2 "Generating Calculations ...", 3
        fOK = CreateAccordExpressionSPs(pfRefreshDatabase)
        gobjProgress.UpdateProgress2
        fOK = fOK And Not gobjProgress.Cancelled
      End If
       
      'Create Payroll purge trigger
      If fOK Then
        OutputCurrentProcess2 "Generating Purge Trigger ..."
        fOK = CreateAccordTransferTriggers(pfRefreshDatabase)
        gobjProgress.UpdateProgress2
        fOK = fOK And Not gobjProgress.Cancelled
      End If
      ' Generate stored procedures for the manual export
      If fOK Then
        OutputCurrentProcess2 "Generating Transfer Procedures ..."
        fOK = CreateAccordTransferSPs(pfRefreshDatabase)
        gobjProgress.UpdateProgress2
        fOK = fOK And Not gobjProgress.Cancelled
      End If
         
    Else
      
      ' Kill any existing triggers
      DropAccordTransferTriggers
      gobjProgress.UpdateProgress False
    
    End If

    ' Save all Relation definitons (MsgBoxErr Done)
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Saving Relations"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveRelations
      fOK = fOK And Not gobjProgress.Cancelled
    End If
  
    'MH20040323
    ' Save the new of modified Outlook Calendar definitions.
    ' Delete the deleted ones.
    If fOK Then
      OutputCurrentProcess "Saving Outlook Calendar Definitions"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveOutlookFolders
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' MH20000728
    ' Save the new of modified Email definitions.
    ' Delete the deleted ones.
    If fOK Then
      OutputCurrentProcess "Saving Email Definitions"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveEmailAddrs(mfrmUse)
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' MH20000731
    ' Create the Email Addr stored procedure.
    If fOK Then
      OutputCurrentProcess "Saving Email Addresses"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = CreateEmailAddrStoredProcedure
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    
    ' Save the new or modified Screen definitions.
    ' Delete the deleted ones.
    If fOK Then
      OutputCurrentProcess "Saving Screen Definitions"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveScreens
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' Save all History Screen definitions.
    If fOK Then
      OutputCurrentProcess "Saving History Screen Definitions"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveHistoryScreens
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' Save the new or modified Workflow definitions.
    ' Delete the deleted ones.
    If fOK Then
      OutputCurrentProcess "Saving Workflow Definitions"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveWorkflows
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' Save the new of modified Order definitions.
    ' Delete the deleted ones.
    If fOK Then
      OutputCurrentProcess "Saving Order Definitions"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveOrders(mfrmUse)
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' Save the new and modified Picture records.
    ' Delete the deleted ones.
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Saving Pictures"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SavePictures
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' Mobile navigation definitions
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Mobile Navigation"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveMobileNavigation
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' Hierarchy specifics.
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Configuring Hierarchy Specifics"
      fOK = ConfigureHierarchySpecifics
      fOK = fOK And Not gobjProgress.Cancelled
    End If
    
    ' Save the new and modified Expressions records.
    ' Delete the deleted ones.
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Saving Calculation Definitions"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveExpressions(pfRefreshDatabase)
      
      'Do progress bar for new scripting bit tacked on the end here
      gobjProgress.ResetBar2
      gobjProgress.Bar2MaxValue = 100
      gobjProgress.Bar2Value = 35
      gobjProgress.UpdateProgress2
      OutputCurrentProcess2 "Scripting System Framework Object"
      
      gobjHRProEngine.Script.ScriptObjects
      
      gobjProgress.Bar2Value = 99
      gobjProgress.UpdateProgress2
      
      fOK = fOK And Not gobjProgress.Cancelled
    End If
     
    ' Create the Overnight SQL Server Job &
    ' Create the Date Dependent column refreshing jobs.
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Generating Overnight Job Schedule"
      gobjProgress.UpdateProgress False
      SaveSystemSetting "overnight", "time", glngOvernightJobTime
      DoEvents
      fOK = CreateOvernightProcess(alngExpressions, pfRefreshDatabase)
      fOK = fOK And Not gobjProgress.Cancelled
    End If
  
    'MH20010403 Need to copy the data BEFORE building the triggers otherwise
    'you will get loads of emails
    ' Copy data from one table to another table (if one was copied from the other).
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Copying Data"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = CopyData
      fOK = fOK And Not gobjProgress.Cancelled
    End If
  
    ' Update OLE field structure
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Upgrading OLE Structures"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = UpgradeOLEDataToV2
      fOK = fOK And Not gobjProgress.Cancelled
    End If
  
    ' Create the Record Validation stored procedures.
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Generating Record Validation"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = CreateValidationStoredProcedures(pfRefreshDatabase)
      fOK = fOK And Not gobjProgress.Cancelled
    End If
  
    ' Create the Column Calculation, Audit and Relationship Triggers.
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Generating Column Triggers"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = fOK And SetTriggers(alngExpressions, pfRefreshDatabase)
      fOK = fOK And gobjHRProEngine.Script.ScriptTriggers
      fOK = fOK And Not gobjProgress.Cancelled
    End If
  
    ' Save any new or modified View definitions.
    ' Delete the deleted ones.
    If fOK Then
      OutputCurrentProcess "Saving View Definitions"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = SaveViews(pfRefreshDatabase)
      fOK = fOK And Not gobjProgress.Cancelled
    End If
     
    ' Save the OLE field stored procedures
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Saving OLE Columns"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = CreateOLEStoredProcedure
      fOK = fOK And Not gobjProgress.Cancelled
    End If
       
    ' Apply permissions
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Applying Permissions"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = ApplyPermissions
      fOK = fOK And Not gobjProgress.Cancelled
    End If
     
    ' Save and check module specifics and configure any stored procedures.
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Module Specifics"
      gobjProgress.UpdateProgress False
      DoEvents
      
      OutputCurrentProcess2 "Saving Definition", 2
      fOK = SaveModuleDefinitions
      fOK = fOK And Not gobjProgress.Cancelled
      gobjProgress.UpdateProgress2
      
      If fOK Then
        fOK = CreateLinkDocumentSP
      End If
      
      If fOK Then
        OutputCurrentProcess2 "Configuring Procedures", 2
        fOK = ConfigureModuleSpecifics
        fOK = fOK And gobjHRProEngine.Script.ScriptFunctions
        fOK = fOK And gobjHRProEngine.Script.ScriptIndexes
        fOK = fOK And Not gobjProgress.Cancelled
        gobjProgress.UpdateProgress2
      End If
      
    End If
   
    ' Save top record of each user table to try and make the first record save a little quicker
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Enabling System Cache"
      gobjProgress.UpdateProgress False
      DoEvents
      fOK = RunRecordSaveOptimiser
      fOK = fOK And Not gobjProgress.Cancelled
    End If
  
  
    ' Reset the 'refresh stored procedures' flag.
    If fOK Then
      gobjProgress.ResetBar2
      OutputCurrentProcess "Resetting save flags"
      gobjProgress.UpdateProgress False
      DoEvents
  
      ' AE20080219 Fault #12905
      If glngEmailMethod = 2 Then
        If Not ValidDatabaseMailDetails(gstrEmailProfile) Then
          glngEmailMethod = 0
          gstrEmailProfile = "<Disable Emails>"
          gstrEmailServer = vbNullString
          gstrEmailAccount = vbNullString
        End If
      End If
      
      SaveSystemSetting "Email", "Method", glngEmailMethod
      SaveSystemSetting "Email", "Profile", gstrEmailProfile
      SaveSystemSetting "Email", "Server", gstrEmailServer
      SaveSystemSetting "Email", "Account", gstrEmailAccount
      
      SaveSystemSetting "Email", "Date Format", glngEmailDateFormat
      SaveSystemSetting "Email", "Attachment Path", gstrEmailAttachmentPath
      SaveSystemSetting "Email", "Test Messages", Replace(gstrEmailTestAddr, "'", "''")
      SaveSystemSetting "Email", "Test on Login", Replace(gstrEmailTestAddr, "'", "''")
  
      'MH20040301
      SaveSystemSetting "Outlook", "AMStartTime", Format(glngAMStartTime, "00:00")
      SaveSystemSetting "Outlook", "AMEndTime", Format(glngAMEndTime, "00:00")
      SaveSystemSetting "Outlook", "PMStartTime", Format(glngPMStartTime, "00:00")
      SaveSystemSetting "Outlook", "PMEndTime", Format(glngPMEndTime, "00:00")
  
      SaveSystemSetting "Database", "RefreshStoredProcedures", 0
      SaveSystemSetting "Database", "UpdatingDateDependantColumns", 0
      SaveSystemSetting "Database", "SystemLastSaveDate", Format(Now, "dd/mm/yyyy")
  
      SaveSystemSetting "ScreenDesigner", "FontName", gobjDefaultScreenFont.Name
      SaveSystemSetting "ScreenDesigner", "FontSize", gobjDefaultScreenFont.Size
      SaveSystemSetting "ScreenDesigner", "FontBold", gobjDefaultScreenFont.Bold
      SaveSystemSetting "ScreenDesigner", "FontItalic", gobjDefaultScreenFont.Italic
      SaveSystemSetting "ScreenDesigner", "FontUnderline", gobjDefaultScreenFont.Underline
      SaveSystemSetting "ScreenDesigner", "FontStrikethrough", gobjDefaultScreenFont.Strikethrough
      SaveSystemSetting "ScreenDesigner", "ForeColor", glngDefaultScreenForeColor
   
    End If
  
    If fOK Then
      SaveSystemSetting "DesktopSetting", "BitmapID", glngDesktopBitmapID
      SaveSystemSetting "DesktopSetting", "BitmapLocation", glngDesktopBitmapLocation
      SaveSystemSetting "DesktopSetting", "BackgroundColour", glngDeskTopColour
    End If
  
    ' Save the audit export (CMG) variables
    If fOK Then
      SaveSystemSetting "CMGExport", "UseCSV", gbCMGExportUseCSV
      SaveSystemSetting "CMGExport", "IgnoreBlanks", gbCMGIgnoreBlanks
      SaveSystemSetting "CMGExport", "ReverseOutput", gbCMGReverseDateChanged
      SaveSystemSetting "CMGExport", "FileCode", gbCMGExportFileCode
      SaveSystemSetting "CMGExport", "FieldCode", gbCMGExportFieldCode
      SaveSystemSetting "CMGExport", "LastChange", gbCMGExportLastChangeDate
      SaveSystemSetting "CMGExport", "FileCodeSize", giCMGExportFileCodeSize
      SaveSystemSetting "CMGExport", "RecordIdentifierSize", giCMGEXportRecordIDSize
      SaveSystemSetting "CMGExport", "FieldCodeSize", giCMGExportFieldCodeSize
      SaveSystemSetting "CMGExport", "OutputColumnSize", giCMGExportOutputColumnSize
      SaveSystemSetting "CMGExport", "LastChangeSize", giCMGExportLastChangeDateSize
      'NPG20090313 Fault 13595
      SaveSystemSetting "CMGExport", "FileCodeOrderID", giCMGExportFileCodeOrderID
      SaveSystemSetting "CMGExport", "RecordIdentifierOrderID", giCMGEXportRecordIDOrderID
      SaveSystemSetting "CMGExport", "FieldCodeOrderID", giCMGExportFieldCodeOrderID
      SaveSystemSetting "CMGExport", "OutputColumnOrderID", giCMGExportOutputColumnOrderID
      SaveSystemSetting "CMGExport", "LastChangeOrderID", giCMGExportLastChangeDateOrderID
    End If
  
    ' Advanced Database settings
    If fOK Then
      SaveSystemSetting "AdvancedDatabaseSetting", "ManualRecursion", gbManualRecursionLevel
      SaveSystemSetting "AdvancedDatabaseSetting", "ManualRecursionLevel", giManualRecursionLevel
      SaveSystemSetting "AdvancedDatabaseSetting", "SpecialFunctionAutoUpdate", gbDisableSpecialFunctionAutoUpdate
      SaveSystemSetting "AdvancedDatabaseSetting", "UpdateIndexesOvernight", gbReorganiseIndexesInOvernightJob
    End If
  
    ' Save the default expression builder settings
    If fOK Then
       SaveSystemSetting "ExpressionBuilder", "ViewColours", glngExpressionViewColours
       SaveSystemSetting "ExpressionBuilder", "NodeSize", glngExpressionViewNodes
    End If
  
    ' Save general defaults
    If fOK Then
      SaveSystemSetting "General", "MaximizeScreens", gbMaximizeScreens
      SaveSystemSetting "General", "RememberColumnsView", gbRememberDBColumnsView
    
      SaveSystemSetting "Web", "SiteAddress", gstrWebSiteAddress
    
      If gbRememberDBColumnsView Then
        SaveSystemSetting "ShowColumn_DataMan", "Name", gpropShowColumns_DataMgr("Name").value
        SaveSystemSetting "ShowColumn_DataMan", "Data Type", gpropShowColumns_DataMgr("Data Type").value
        SaveSystemSetting "ShowColumn_DataMan", "Size", gpropShowColumns_DataMgr("Size").value
        SaveSystemSetting "ShowColumn_DataMan", "Decimals", gpropShowColumns_DataMgr("Decimals").value
        SaveSystemSetting "ShowColumn_DataMan", "Display", gpropShowColumns_DataMgr("Display").value
        SaveSystemSetting "ShowColumn_DataMan", "Column Type", gpropShowColumns_DataMgr("Column Type").value
        SaveSystemSetting "ShowColumn_DataMan", "Control Type", gpropShowColumns_DataMgr("Control Type").value
        SaveSystemSetting "ShowColumn_DataMan", "Read Only", gpropShowColumns_DataMgr("Read Only").value
        SaveSystemSetting "ShowColumn_DataMan", "Audit", gpropShowColumns_DataMgr("Audit").value
        SaveSystemSetting "ShowColumn_DataMan", "Multi-line", gpropShowColumns_DataMgr("Multi-line").value
        SaveSystemSetting "ShowColumn_DataMan", "Case", gpropShowColumns_DataMgr("Case").value
        SaveSystemSetting "ShowColumn_DataMan", "Default Value", gpropShowColumns_DataMgr("Default Value").value
        SaveSystemSetting "ShowColumn_DataMan", "Text Alignment", gpropShowColumns_DataMgr("Text Alignment").value
        SaveSystemSetting "ShowColumn_DataMan", "Duplicate Check", gpropShowColumns_DataMgr("Duplicate Check").value
        SaveSystemSetting "ShowColumn_DataMan", "Mandatory", gpropShowColumns_DataMgr("Mandatory").value
        SaveSystemSetting "ShowColumn_DataMan", "Unique in Table", gpropShowColumns_DataMgr("Unique in Table").value
        SaveSystemSetting "ShowColumn_DataMan", "Unique in Siblings", gpropShowColumns_DataMgr("Unique in Siblings").value
        SaveSystemSetting "ShowColumn_DataMan", "Mask", gpropShowColumns_DataMgr("Mask").value
        SaveSystemSetting "ShowColumn_DataMan", "Custom Validation", gpropShowColumns_DataMgr("Custom Validation").value
        SaveSystemSetting "ShowColumn_DataMan", "Diary Links", gpropShowColumns_DataMgr("Diary Links").value
        'SaveSystemSetting "ShowColumn_DataMan", "Email Links", gpropShowColumns_DataMgr("Email Links").value
        
        ' JDM - Fault 6545 - Was causing error in systems where AFD is not enabled.
        If IsModuleEnabled(modAFD) Then
          SaveSystemSetting "ShowColumn_DataMan", "AFD Postcode", gpropShowColumns_DataMgr("AFD Postcode").value
        End If
        
        If IsModuleEnabled(modQAddress) Then
          SaveSystemSetting "ShowColumn_DataMan", "Quick Address", gpropShowColumns_DataMgr("Quick Address").value
        End If
        
        'NHRD29072003 Fault 6208 Added the ability to show and store Use1000Separator and
        'Trimming column properties
        SaveSystemSetting "ShowColumn_DataMan", "Use 1000 Separator", gpropShowColumns_DataMgr("Use 1000 Separator").value
        SaveSystemSetting "ShowColumn_DataMan", "Trimming", gpropShowColumns_DataMgr("Trimming").value
        
        SaveSystemSetting "ShowColumn_DataManTable", "Name", gpropShowColumns_DataMgrTable("Name").value
        SaveSystemSetting "ShowColumn_DataManTable", "Type", gpropShowColumns_DataMgrTable("Type").value
        SaveSystemSetting "ShowColumn_DataManTable", "Primary Order", gpropShowColumns_DataMgrTable("Primary Order").value
        SaveSystemSetting "ShowColumn_DataManTable", "Record Description", gpropShowColumns_DataMgrTable("Record Description").value
        SaveSystemSetting "ShowColumn_DataManTable", "Default Email", gpropShowColumns_DataMgrTable("Default Email").value
        SaveSystemSetting "ShowColumn_DataManTable", "Email Links", gpropShowColumns_DataMgrTable("Email Links").value
        SaveSystemSetting "ShowColumn_DataManTable", "Calendar Links", gpropShowColumns_DataMgrTable("Calendar Links").value
        
        If IsModuleEnabled(modWorkflow) Then
          SaveSystemSetting "ShowColumn_DataManTable", "Workflow Links", gpropShowColumns_DataMgrTable("Workflow Links").value
        End If
        
        SaveSystemSetting "ShowColumn_PictMan", "Name", gpropShowColumns_PictMgr("Name").value
        SaveSystemSetting "ShowColumn_PictMan", "Type", gpropShowColumns_PictMgr("Type").value
        SaveSystemSetting "ShowColumn_PictMan", "Height", gpropShowColumns_PictMgr("Height").value
        SaveSystemSetting "ShowColumn_PictMan", "Width", gpropShowColumns_PictMgr("Width").value
        
        SaveSystemSetting "ShowColumn_ViewMan", "Name", gpropShowColumns_ViewMgr("Name").value
        SaveSystemSetting "ShowColumn_ViewMan", "Description", gpropShowColumns_ViewMgr("Description").value
        
      End If
    End If
  
    ' SQL Process Info
    SaveSystemSetting "ProcessAccount", "Mode", glngProcessMethod
  
    ' Save the current configuration of the platform
    If fOK Then
      SaveSystemSetting "Platform", "SQLServerVersion", gstrSQLFullVersion
      ' AE20080215 - Changed to get the actual Server/DB name
      SaveSystemSetting "Platform", "DatabaseName", IIf(GetDBName() = vbNullString, gsDatabaseName, GetDBName())
      SaveSystemSetting "Platform", "ServerName", IIf(GetServerName() = vbNullString, gsServerName, GetServerName())
    End If
  
    ' Apply any post save hotfixes
    If fOK Then
      fOK = ApplyHotfixes(AFTERSAVE)
    End If
  
    If fOK Then
      gobjProgress.UpdateProgress False
      DoEvents
    End If

  End If
  
TidyUpAndExit:
  
  OutputCurrentProcess vbNullString
  
  If Not gobjHRProEngine.ErrorLog Is Nothing Then
    gobjHRProEngine.ErrorLog.OutputToFile (gsLogDirectory + "\OpenHRFramework.log")
    If gobjHRProEngine.ErrorLog.ErrorCount > 0 Then
      gobjProgress.Visible = False
      gobjHRProEngine.ErrorLog.Show
      fOK = Not gobjHRProEngine.ErrorLog.IsCatastrophic
    End If
    
  End If
  
  If fOK Then
    
    AuditAccess "Save", "System"
    
    ' Commit transactions
    OutputCurrentProcess "Committing changes to the server"
    
    ' Close the .NET databases
    'objPhoenix.CommitDB.CommitTrans
   ' objPhoenix.MetadataDB.Close
    
    gADOCon.CommitTrans
   
    ' Apply databsae ownership as required.
    gobjProgress.ResetBar2
    OutputCurrentProcess "Applying Database Ownership"
    gobjProgress.UpdateProgress False
    DoEvents
    ApplyDatabaseOwnership
    ApplyPostSaveProcessing
    
    gobjProgress.UpdateProgress False
    DoEvents
    
    'MH20010302 Fault 1228
    'Needs to be done before CreateTempTables as we need to check which tables have changed
    CheckIfRebuildDiaryOrEmail

    '27/07/2001 MH
    If Application.AccessMode = accSupportMode Then
      Application.AccessMode = accLimited
      frmSysMgr.SetCaption
    End If

    '16/08/2001 MH Fault 2691
    gfRefreshStoredProcedures = False

    OutputCurrentProcess "Reinitialising Interface"
    gobjProgress.UpdateProgress False
    DoEvents


    ' Refresh the local tables.
    fOK = DropTempTables
    If fOK Then
      fOK = CreateTempTables
    End If
  
  Else
    ' Rollback transactions.
    If fInTransaction Then
      
      'MH20010329 Put the next two lines of code in if you want the progress bar to be nearly complete.
      OutputCurrentProcess "Undoing changes to the server database"
      gobjProgress.Visible = True

      Screen.MousePointer = vbHourglass
   '   objPhoenix.CommitDB.RollbackTrans
      gADOCon.RollbackTrans
      Screen.MousePointer = vbDefault
    End If
  End If

  If bLockedOK Then
    UnlockDatabase (lckSaving)
  End If
  
  ' Clear wait message.
  If Forms.Count > 0 Then
    frmSysMgr.StatusBar1.SimpleText = gsDatabaseName & _
      " - Version : " & gstrSQLFullVersion
    frmSysMgr.StatusBar1.Refresh
  End If
  Screen.MousePointer = vbDefault
  
  SaveChanges = fOK
  
  gobjProgress.CloseProgress
  
  'Re-enable all controls
  EnableOpenForms True

  frmSysMgr.tmrKeepAlive.Enabled = True

  OutputCurrentProcess "End of save process"
  Exit Function

ErrorTrap:
  fOK = False
  gobjProgress.ResetBar2
  OutputError "Error saving changes"
  Resume TidyUpAndExit
  
End Function


Public Function SaveChanges_LogoutCheck(blnSendMessageVisible As Boolean) As Boolean

  Dim frmViewUsers As frmViewCurrentUsers
  Dim blnCancelled As Boolean
  Dim fOK As Boolean
  
  Set frmViewUsers = New frmViewCurrentUsers
  With frmViewUsers

    fOK = .OkayToSave
    If Not fOK And .grdUsers.Rows > 0 Then

      gobjProgress.Visible = False
      Screen.MousePointer = vbDefault
      
      'NHRD20030425 Fault 4880 Added a check to determine 'which direction' theyre coming from.
      'i.e. if they are doing a quick update then they wouldn't be logged in yet. Conversely if
      'they are saving changes then they would already be logged in and I can engineer the message accordingly.
      If Application.LoggedIn() Then
        MsgBox "Making changes to the database will affect users who are currently logged into the system." & vbNewLine & vbNewLine & _
               "You will need to ensure that all users are logged out and that you have locked the system " & _
               "before you can apply these changes.", vbInformation, "Saving Changes"
      Else
        MsgBox "Updating the system will affect users who are currently logged in." & vbNewLine & vbNewLine & _
               "You will need to ensure that all users are logged out " & _
               "before you can run the update process.", vbInformation, "Updating System"
      End If
      
      If .OkayToSave = False Then
        .Enabled = True
        .Saving = True
        .SendMessageVisible = blnSendMessageVisible
        .Show vbModal
      End If
      Screen.MousePointer = vbHourglass
      fOK = Not .Cancelled
    
      If .ForciblyDisconnect Then
        AuditAccess "Forcibly Disconnect Users", "System"
      End If
    
      ' AE20080422 Fault #13121
      If fOK Then
        gobjProgress.Visible = True
      Else
        ' TM20080112 Fault #13391 - remove the save lock if cancelled out of View Current Users dialog.
        UnlockDatabase (lckSaving)
      End If
      
    ' AE20080325 Fault #12903
    ElseIf .Locked Then
      gobjProgress.Visible = False
      Call UpdateLockCheck
      fOK = False
    End If

    SaveChanges_LogoutCheck = fOK

  End With
  UnLoad frmViewUsers
  Set frmViewUsers = Nothing

End Function


Private Function SaveModuleDefinitions() As Boolean
  ' Save the module definitions.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim rsModules As New ADODB.Recordset
  Dim rsRelatedColumns As New ADODB.Recordset
  Dim rsLinks As DAO.Recordset
  Dim rsAccord As DAO.Recordset
'  Dim rsFusion As DAO.Recordset
  Dim rsData As DAO.Recordset
  Dim sSQL As String
  Dim alngLinkIDs() As Long
  Dim rsMaxLinkID As New ADODB.Recordset
  Dim iLoop As Integer
  Dim lngNewLinkID As Long
  
  fOK = True
  
  ' Default some Workflow setup parameters
  DefaultWorkflowSetup
  
  ' Delete any existing Module definitions.
  gADOCon.Execute "DELETE FROM ASRSysModuleSetup", , adCmdText + adExecuteNoRecords
  
  ' Open the Module Setup table.
  rsModules.Open "ASRSysModuleSetup", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
  
  With recModuleSetup
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      rsModules.AddNew
      rsModules!moduleKey = !moduleKey
      rsModules!parameterkey = !parameterkey
      
      If Not IsNull(!parametervalue) Then
        rsModules!parametervalue = !parametervalue
      End If
      
      rsModules!ParameterType = !ParameterType
      rsModules.Update
      
      .MoveNext
    Loop
  End With
  
  rsModules.Close
  
  ' Delete any existing Related Column definitions.
  gADOCon.Execute "DELETE FROM ASRSysModuleRelatedColumns", , adCmdText + adExecuteNoRecords
  
  ' Open the Module Related Column table.
  rsRelatedColumns.Open "ASRSysModuleRelatedColumns", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
  
  With recModuleRelatedColumns
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      rsRelatedColumns.AddNew
      rsRelatedColumns!moduleKey = !moduleKey
      rsRelatedColumns!parameterkey = !parameterkey
      rsRelatedColumns!sourcecolumnid = !sourcecolumnid
      rsRelatedColumns!destcolumnid = !destcolumnid
      rsRelatedColumns.Update
      
      .MoveNext
    Loop
  End With
  
  rsRelatedColumns.Close

  ' Delete any existing Self-service Intranet Link definitions.
  gADOCon.Execute "DELETE FROM ASRSysSSIntranetLinks", , adCmdText + adExecuteNoRecords

  ReDim alngLinkIDs(2, 0)
  
  sSQL = "SELECT *" & _
    " FROM tmpSSIntranetLinks"
  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  While Not rsLinks.EOF
    'JPD 20040630 Fault 8859
    'sSQL = "INSERT INTO ASRSysSSIntranetLinks" & _
      " (linkType, linkOrder, prompt, text, screenID, pageTitle, URL, startMode, utilityType, utilityID, viewID)" & _
      " VALUES(" & _
      CStr(rsLinks!LinkType) & "," & _
      CStr(rsLinks!linkOrder) & "," & _
      "'" & Replace(rsLinks!Prompt, "'", "''") & "'," & _
      "'" & Replace(rsLinks!Text, "'", "''") & "'," & _
      CStr(rsLinks!ScreenID) & "," & _
      "'" & Replace(rsLinks!PageTitle, "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!URL), "", rsLinks!URL), "'", "''") & "'," & _
      CStr(rsLinks!StartMode) & "," & _
      CStr(IIf(IsNull(rsLinks!UtilityType), 0, rsLinks!UtilityType)) & "," & _
      CStr(IIf(IsNull(rsLinks!UtilityID), 0, rsLinks!UtilityID)) & "," & _
      CStr(IIf(IsNull(rsLinks!ViewID), 0, rsLinks!ViewID)) & _
      ")"
      
    'NHRD31012007 Open in New Window Development ammendment.
    'Added the newWindow variable
    'NPG20080125 Fault 12873 Added EMailAddress & EMail Subject
    ' NPG20100126 SSI Dashboard elements added
    sSQL = "INSERT INTO ASRSysSSIntranetLinks" & _
      " (linkType, linkOrder, prompt, text, screenID, pageTitle, URL, startMode, utilityType, utilityID, " & _
      "viewID, newWindow, tableID, EMailAddress, EMailSubject, AppFilePath, AppParameters, " & _
      "DocumentFilePath, DisplayDocumentHyperlink, Element_Type, SeparatorOrientation, PictureID, Chart_ShowLegend, " & _
      "Chart_Type, Chart_ShowGrid, Chart_StackSeries, Chart_viewID, Chart_TableID, Chart_ColumnID, Chart_FilterID, " & _
      "Chart_AggregateType, Chart_ShowValues, UseFormatting, Formatting_DecimalPlaces, Formatting_Use1000Separator, " & _
      "Formatting_Prefix, Formatting_Suffix, UseConditionalFormatting, ConditionalFormatting_Operator_1, ConditionalFormatting_Value_1, " & _
      "ConditionalFormatting_Style_1, ConditionalFormatting_Colour_1, ConditionalFormatting_Operator_2, ConditionalFormatting_Value_2, " & _
      "ConditionalFormatting_Style_2, ConditionalFormatting_Colour_2, ConditionalFormatting_Operator_3, ConditionalFormatting_Value_3, " & _
      "ConditionalFormatting_Style_3, ConditionalFormatting_Colour_3, SeparatorColour, InitialDisplayMode, Chart_TableID_2, Chart_ColumnID_2, " & _
      "Chart_TableID_3, Chart_ColumnID_3, Chart_SortOrderID, Chart_SortDirection, Chart_ColourID, Chart_ShowPercentages)" & _
      " VALUES(" & _
      CStr(rsLinks!LinkType) & "," & _
      CStr(rsLinks!linkOrder) & "," & _
      "'" & Replace(rsLinks!Prompt, "'", "''") & "'," & _
      "'" & Replace(rsLinks!Text, "'", "''") & "'," & _
      CStr(rsLinks!ScreenID) & "," & _
      "'" & Replace(rsLinks!PageTitle, "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!URL), "", rsLinks!URL), "'", "''") & "',"
    sSQL = sSQL & _
      CStr(rsLinks!StartMode) & "," & _
      CStr(IIf(IsNull(rsLinks!UtilityType), 0, rsLinks!UtilityType)) & "," & _
      CStr(IIf(IsNull(rsLinks!UtilityID), 0, rsLinks!UtilityID)) & "," & _
      CStr(IIf(IsNull(rsLinks!ViewID), 0, rsLinks!ViewID)) & "," & _
      IIf(IsNull(rsLinks!NewWindow), "0", IIf(rsLinks!NewWindow, "1", "0")) & "," & _
      CStr(IIf(IsNull(rsLinks!TableID), 0, rsLinks!TableID)) & "," & _
      "'" & Replace(IIf(IsNull(rsLinks!EMailAddress), "", rsLinks!EMailAddress), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!EMailSubject), "", rsLinks!EMailSubject), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!AppFilePath), "", rsLinks!AppFilePath), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!AppParameters), "", rsLinks!AppParameters), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!DocumentFilePath), "", rsLinks!DocumentFilePath), "'", "''") & "'," & _
      IIf(IsNull(rsLinks!DisplayDocumentHyperlink), "0", IIf(rsLinks!DisplayDocumentHyperlink, "1", "0")) & "," & _
      IIf(IsNull(rsLinks!Element_Type), "0", rsLinks!Element_Type) & ","
    sSQL = sSQL & _
      IIf(IsNull(rsLinks!SeparatorOrientation), 0, rsLinks!SeparatorOrientation) & "," & _
      IIf(IsNull(rsLinks!PictureID), 0, rsLinks!PictureID) & "," & _
      IIf(IsNull(rsLinks!Chart_ShowLegend), "0", IIf(rsLinks!Chart_ShowLegend, "1", "0")) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_Type), 0, rsLinks!Chart_Type)) & "," & _
      IIf(IsNull(rsLinks!Chart_ShowGrid), "0", IIf(rsLinks!Chart_ShowGrid, "1", "0")) & "," & _
      IIf(IsNull(rsLinks!Chart_StackSeries), "0", IIf(rsLinks!Chart_StackSeries, "1", "0")) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_ViewID), 0, rsLinks!Chart_ViewID)) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_TableID), 0, rsLinks!Chart_TableID)) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_ColumnID), 0, rsLinks!Chart_ColumnID)) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_FilterID), 0, rsLinks!Chart_FilterID)) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_AggregateType), 0, rsLinks!Chart_AggregateType)) & ","

    sSQL = sSQL & _
      IIf(IsNull(rsLinks!Chart_ShowValues), "0", IIf(rsLinks!Chart_ShowValues, "1", "0")) & "," & _
      IIf(IsNull(rsLinks!UseFormatting), "0", IIf(rsLinks!UseFormatting, "1", "0")) & "," & _
      CStr(IIf(IsNull(rsLinks!Formatting_DecimalPlaces), 0, rsLinks!Formatting_DecimalPlaces)) & "," & _
      IIf(IsNull(rsLinks!Formatting_Use1000Separator), "0", IIf(rsLinks!Formatting_Use1000Separator, "1", "0")) & "," & _
      "'" & Replace(IIf(IsNull(rsLinks!Formatting_Prefix), "", rsLinks!Formatting_Prefix), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!Formatting_Suffix), "", rsLinks!Formatting_Suffix), "'", "''") & "'," & _
      IIf(IsNull(rsLinks!UseConditionalFormatting), "0", IIf(rsLinks!UseConditionalFormatting, "1", "0")) & "," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Operator_1), "", rsLinks!ConditionalFormatting_Operator_1), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Value_1), "", rsLinks!ConditionalFormatting_Value_1), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Style_1), "", rsLinks!ConditionalFormatting_Style_1), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Colour_1), "", rsLinks!ConditionalFormatting_Colour_1), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Operator_2), "", rsLinks!ConditionalFormatting_Operator_2), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Value_2), "", rsLinks!ConditionalFormatting_Value_2), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Style_2), "", rsLinks!ConditionalFormatting_Style_2), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Colour_2), "", rsLinks!ConditionalFormatting_Colour_2), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Operator_3), "", rsLinks!ConditionalFormatting_Operator_3), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Value_3), "", rsLinks!ConditionalFormatting_Value_3), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Style_3), "", rsLinks!ConditionalFormatting_Style_3), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!ConditionalFormatting_Colour_3), "", rsLinks!ConditionalFormatting_Colour_3), "'", "''") & "'," & _
      "'" & Replace(IIf(IsNull(rsLinks!SeparatorColour), "", rsLinks!SeparatorColour), "'", "''") & "'," & CStr(IIf(IsNull(rsLinks!InitialDisplayMode), 0, rsLinks!InitialDisplayMode)) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_TableID_2), 0, rsLinks!Chart_TableID_2)) & "," & CStr(IIf(IsNull(rsLinks!Chart_ColumnID_2), 0, rsLinks!Chart_ColumnID_2)) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_TableID_3), 0, rsLinks!Chart_TableID_3)) & "," & CStr(IIf(IsNull(rsLinks!Chart_ColumnID_3), 0, rsLinks!Chart_ColumnID_3)) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_SortOrderID), 0, rsLinks!Chart_SortOrderID)) & "," & CStr(IIf(IsNull(rsLinks!Chart_SortDirection), 0, rsLinks!Chart_SortDirection)) & "," & _
      CStr(IIf(IsNull(rsLinks!Chart_ColourID), 0, rsLinks!Chart_ColourID)) & "," & IIf(IsNull(rsLinks!Chart_ShowPercentages), "0", IIf(rsLinks!Chart_ShowPercentages, "1", "0")) & ")"

    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    sSQL = "SELECT MAX(id) AS [result]" & _
      " FROM ASRSysSSIntranetLinks"
    rsMaxLinkID.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    ReDim Preserve alngLinkIDs(2, UBound(alngLinkIDs, 2) + 1)
    alngLinkIDs(1, UBound(alngLinkIDs, 2)) = rsLinks!ID
    alngLinkIDs(2, UBound(alngLinkIDs, 2)) = rsMaxLinkID!result
    rsMaxLinkID.Close
    
    rsLinks.MoveNext
  Wend
  rsLinks.Close

  ' Delete any existing Self-service Intranet Link Hidden Group records.
  gADOCon.Execute "DELETE FROM ASRSysSSIHiddenGroups", , adCmdText + adExecuteNoRecords

  sSQL = "SELECT *" & _
    " FROM tmpSSIHiddenGroups"
  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  While Not rsLinks.EOF
    lngNewLinkID = 0
    For iLoop = 1 To UBound(alngLinkIDs, 2)
      If alngLinkIDs(1, iLoop) = rsLinks!LinkID Then
        lngNewLinkID = alngLinkIDs(2, iLoop)
        Exit For
      End If
    Next iLoop
    
    sSQL = "INSERT INTO ASRSysSSIHiddenGroups" & _
      " (linkID, groupName)" & _
      " VALUES(" & _
      CStr(lngNewLinkID) & "," & _
      "'" & Replace(rsLinks!GroupName, "'", "''") & "'" & _
      ")"

    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    rsLinks.MoveNext
  Wend
  rsLinks.Close

  ' Delete any existing Self-service Intranet Views records.
  gADOCon.Execute "DELETE FROM ASRSysSSIViews", , adCmdText + adExecuteNoRecords

  sSQL = "SELECT *" & _
    " FROM tmpSSIViews"
  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  While Not rsLinks.EOF
      
    sSQL = "INSERT INTO ASRSysSSIViews" & _
      " (viewID, " & _
      "    buttonLinkPromptText, " & _
      "    buttonLinkButtonText, " & _
      "    hypertextLinkText, " & _
      "    dropdownListLinkText, " & _
      "    buttonLink, " & _
      "    hypertextLink, " & _
      "    dropdownListLink, " & _
      "    singleRecordView, " & _
      "    sequence, " & _
      "    linksLinkText, " & _
      "    pageTitle, " & _
      "    tableID, " & _
      "    WFOutOfOffice" & _
      ")"
    sSQL = sSQL & _
      " VALUES(" & _
      CStr(rsLinks!ViewID) & "," & _
      "'" & Replace(rsLinks!ButtonLinkPromptText, "'", "''") & "'," & _
      "'" & Replace(rsLinks!ButtonLinkButtonText, "'", "''") & "'," & _
      "'" & Replace(rsLinks!HypertextLinkText, "'", "''") & "'," & _
      "'" & Replace(rsLinks!DropdownListLinkText, "'", "''") & "'," & _
      IIf(rsLinks!ButtonLink, "1", "0") & "," & _
      IIf(rsLinks!HypertextLink, "1", "0") & "," & _
      IIf(rsLinks!DropdownListLink, "1", "0") & "," & _
      IIf(rsLinks!SingleRecordView, "1", "0") & "," & _
      CStr(rsLinks!Sequence) & "," & _
      "'" & Replace(rsLinks!LinksLinkText, "'", "''") & "'," & _
      "'" & IIf(IsNull(rsLinks!PageTitle), vbNullString, Replace(IIf(IsNull(rsLinks!PageTitle), vbNullString, rsLinks!PageTitle), "'", "''")) & "'," & _
      CStr(rsLinks!TableID) & ", " & IIf(rsLinks!WFOutOfOffice, "1", "0") & _
      ")"

    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    rsLinks.MoveNext
  Wend
  rsLinks.Close

  ' Store the Payroll Transfer Types
  gADOCon.Execute "DELETE FROM ASRSysAccordTransferTypes", , adCmdText + adExecuteNoRecords
  
  sSQL = "SELECT * FROM tmpAccordTransferTypes"
  Set rsAccord = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  While Not rsAccord.EOF
    
    sSQL = "INSERT INTO ASRSysAccordTransferTypes" & _
      " (TransferTypeID, TransferType, FilterID, ASRBaseTableID, IsVisible, ForceAsUpdate)" & _
      " VALUES (" & _
      CStr(rsAccord!TransferTypeID) & "," & _
      "'" & CStr(rsAccord!TransferType) & "'," & _
      CStr(rsAccord!FilterID) & "," & _
      CStr(rsAccord!ASRBaseTableID) & "," & _
      IIf(rsAccord!IsVisible, "1", "0") & "," & _
      IIf(rsAccord!ForceAsUpdate, "1", "0") & ")"
    
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    rsAccord.MoveNext
  Wend
  rsAccord.Close

  ' Store the Payroll mappings
  gADOCon.Execute "DELETE FROM ASRSysAccordTransferFieldDefinitions", , adCmdText + adExecuteNoRecords
  
  sSQL = "SELECT * FROM tmpAccordTransferFieldDefinitions"
  Set rsAccord = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  While Not rsAccord.EOF
    
    sSQL = "INSERT INTO ASRSysAccordTransferFieldDefinitions" & _
      " (TransferFieldID, TransferTypeID, Mandatory, Description, ASRMapType, ASRTableID, ASRColumnID, ASRExprID, ASRValue, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ConvertData" & _
      " , IsEmployeeName, IsDepartmentCode, IsDepartmentName, IsPayrollCode, GroupBy, PreventModify) " & _
      " VALUES (" & _
      CStr(rsAccord!TransferFieldID) & "," & _
      CStr(rsAccord!TransferTypeID) & "," & _
      IIf(rsAccord!Mandatory, "1", "0") & "," & _
      "'" & Replace(rsAccord!Description, "'", "''") & "'," & _
      IIf(IsNull(rsAccord!ASRMapType), "null", rsAccord!ASRMapType) & "," & _
      IIf(IsNull(rsAccord!ASRTableID), "null", rsAccord!ASRTableID) & "," & _
      IIf(IsNull(rsAccord!ASRColumnID), "null", rsAccord!ASRColumnID) & "," & _
      IIf(IsNull(rsAccord!ASRExprID), "null", rsAccord!ASRExprID) & "," & _
      "'" & Replace(IIf(IsNull(rsAccord!ASRValue), vbNullString, rsAccord!ASRValue), "'", "''") & "'," & _
      IIf(rsAccord!IsCompanyCode, "1", "0") & "," & _
      IIf(rsAccord!IsEmployeeCode, "1", "0") & "," & _
      CStr(rsAccord!Direction) & "," & _
      IIf(rsAccord!IsKeyField, "1", "0") & "," & _
      IIf(rsAccord!AlwaysTransfer, "1", "0") & "," & _
      IIf(rsAccord!ConvertData, "1", "0") & "," & _
      IIf(rsAccord!IsEmployeeName, "1", "0") & "," & _
      IIf(rsAccord!IsDepartmentCode, "1", "0") & "," & _
      IIf(rsAccord!IsDepartmentName, "1", "0") & "," & _
      IIf(rsAccord!IsPayrollCode, "1", "0") & "," & _
      IIf(IsNull(rsAccord!GroupBy), "null", rsAccord!GroupBy) & ", " & _
      IIf(rsAccord!PreventModify, "1", "0") & ")"

    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    rsAccord.MoveNext
  Wend
  rsAccord.Close

  ' Store the Payroll Column Value Mappings
  gADOCon.Execute "DELETE FROM ASRSysAccordTransferFieldMappings", , adCmdText + adExecuteNoRecords
  
  sSQL = "SELECT * FROM tmpAccordTransferFieldMappings"
  Set rsAccord = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  While Not rsAccord.EOF

    sSQL = "INSERT INTO ASRSysAccordTransferFieldMappings" & _
      " (TransferID, FieldID, HRProValue, AccordValue)" & _
      " VALUES (" & _
      CStr(rsAccord!TransferID) & "," & _
      CStr(rsAccord!FieldID) & "," & _
      "'" & rsAccord!HRProValue & "'," & _
      "'" & rsAccord!AccordValue & "')"
      
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    rsAccord.MoveNext
  Wend
  rsAccord.Close


'  ' Store the Fusion Transfer Types
'  gADOCon.Execute "DELETE FROM ASRSysFusionTypes", , adCmdText + adExecuteNoRecords
'
'  sSQL = "SELECT * FROM tmpFusionTypes"
'  Set rsFusion = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  While Not rsFusion.EOF
'
'    sSQL = "INSERT INTO ASRSysFusionTypes" & _
'      " (FusionTypeID, FusionType, FilterID, ASRBaseTableID, IsVisible, Version)" & _
'      " VALUES (" & _
'      CStr(rsFusion!FusionTypeID) & "," & _
'      "'" & CStr(rsFusion!FusionType) & "'," & _
'      CStr(rsFusion!FilterID) & "," & _
'      IIf(IsNull(rsFusion!ASRBaseTableID), "0", CStr(rsFusion!ASRBaseTableID)) & "," & _
'      IIf(rsFusion!IsVisible, "1", "0") & "," & _
'      "1" & ")"
'
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'    rsFusion.MoveNext
'  Wend
'  rsFusion.Close
'
'  ' Store the Fusion mappings
'  gADOCon.Execute "DELETE FROM ASRSysFusionFieldDefinitions", , adCmdText + adExecuteNoRecords
'
'  sSQL = "SELECT * FROM tmpFusionFieldDefinitions"
'  Set rsFusion = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  While Not rsFusion.EOF
'
'    sSQL = "INSERT INTO ASRSysFusionFieldDefinitions" & _
'      " (FusionTypeID, Nodekey, Mandatory, Description, ASRMapType, ASRTableID, ASRColumnID, ASRExprID, ASRValue, IsCompanyCode, IsEmployeeCode, IsKeyField, AlwaysTransfer, ConvertData" & _
'      ", IsEmployeeName, IsDepartmentCode, IsDepartmentName, PreventModify, DataType) " & _
'      " VALUES (" & _
'      CStr(rsFusion!FusionTypeID) & ",'" & _
'      CStr(rsFusion!NodeKey) & "'," & _
'      IIf(rsFusion!Mandatory, "1", "0") & "," & _
'      "'" & Replace(rsFusion!Description, "'", "''") & "'," & _
'      IIf(IsNull(rsFusion!ASRMapType), "null", rsFusion!ASRMapType) & "," & _
'      IIf(IsNull(rsFusion!ASRTableID), "null", rsFusion!ASRTableID) & "," & _
'      IIf(IsNull(rsFusion!ASRColumnID), "null", rsFusion!ASRColumnID) & "," & _
'      IIf(IsNull(rsFusion!ASRExprID), "null", rsFusion!ASRExprID) & "," & _
'      "'" & Replace(IIf(IsNull(rsFusion!ASRValue), vbNullString, rsFusion!ASRValue), "'", "''") & "'," & _
'      IIf(rsFusion!IsCompanyCode, "1", "0") & "," & _
'      IIf(rsFusion!IsEmployeeCode, "1", "0") & "," & _
'      "" & _
'      IIf(rsFusion!IsKeyField, "1", "0") & "," & _
'      IIf(rsFusion!AlwaysTransfer, "1", "0") & "," & _
'      IIf(rsFusion!ConvertData, "1", "0") & "," & _
'      IIf(rsFusion!IsEmployeeName, "1", "0") & "," & _
'      IIf(rsFusion!IsDepartmentCode, "1", "0") & _
'      "" & _
'      ",0, 0" & _
'      IIf(rsFusion!PreventModify, "1", "0") & "," & CStr(rsFusion!DataType) & ")"
'
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'    rsFusion.MoveNext
'  Wend
'  rsFusion.Close
'
'  ' Store the Fusion Column Value Mappings
'  gADOCon.Execute "DELETE FROM ASRSysFusionFieldMappings", , adCmdText + adExecuteNoRecords
'
'  sSQL = "SELECT * FROM tmpFusionFieldMappings"
'  Set rsFusion = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'  While Not rsFusion.EOF
'
'    sSQL = "INSERT INTO ASRSysFusionFieldMappings" & _
'      " (TransferID, FieldID, HRProValue, FusionValue)" & _
'      " VALUES (" & _
'      CStr(rsFusion!TransferID) & "," & _
'      CStr(rsFusion!FieldID) & "," & _
'      "'" & rsFusion!HRProValue & "'," & _
'      "'" & rsFusion!FusionValue & "')"
'
'    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
'
'    rsFusion.MoveNext
'  Wend
'  rsFusion.Close

TidyUpAndExit:
  Set rsAccord = Nothing
  'Set rsFusion = Nothing
  Set rsMaxLinkID = Nothing
  Set rsModules = Nothing
  Set rsRelatedColumns = Nothing
  SaveModuleDefinitions = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error saving module setup"
  Resume TidyUpAndExit

End Function


Private Function SaveRelations() As Boolean
  ' Save the Relation definition to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim rsRelations As ADODB.Recordset
  
  Set rsRelations = New ADODB.Recordset
  fOK = True
  
  ' Delete any existing relation definitions.
  gADOCon.Execute "DELETE FROM ASRSysRelations", , adCmdText + adExecuteNoRecords
  
  ' Open the relations table.
  rsRelations.Open "ASRSysRelations", gADOCon, adOpenDynamic, adLockOptimistic, adCmdTableDirect
  
  With recRelEdit
    ' Loop through relations table in local database
    ' add to relations table in remote database.
    If Not (.EOF And .BOF) Then
      .MoveFirst
      Do While Not .EOF
        rsRelations.AddNew
        rsRelations!parentID = !parentID
        rsRelations!childID = !childID
        rsRelations.Update
        
        .MoveNext
      Loop
    End If
  End With
  
  rsRelations.Close
  
TidyUpAndExit:
  Set rsRelations = Nothing
  SaveRelations = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  'gobjProgress.Visible = False
  'MsgBox ODBC.FormatError(Err.Description), vbOKOnly, vbExclamation, Application.Name
  OutputError "Error saving relations"
  Resume TidyUpAndExit

End Function


Private Function CleanupDatabase() As Boolean

  ' Clean any temporary stored procedures/functions/udfs that are lying around in the database
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  ' Delete the existing order definition from the server database.
  sSQL = "EXEC spASRDropTempObjects"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  ' Clear out any junk that may be laying round in the messages table
  sSQL = "DELETE FROM ASRSysMessages"
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  CleanupDatabase = fOK
  Exit Function

ErrorTrap:
  fOK = False
  OutputError "Error cleaning database"
  Resume TidyUpAndExit

End Function


Private Function CopyData() As Boolean
  ' Copy the data to any cloned tables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iSourceColumnDataType As Integer
  Dim iDestinationColumnSize As Integer
  Dim iDestinationColumnDecimals As Integer
  Dim iDestinationColumnDataType As Integer
  Dim lngSourceTableID As Long
  Dim lngDestinationTableID As Long
  Dim dblMaxValue As Double
  Dim sSQL As String
  Dim sTempCopy As String
  Dim sValueList As SystemMgr.cStringBuilder
  Dim sColumnList As SystemMgr.cStringBuilder
  Dim sSourceTableName As String
  Dim sDestinationTableName As String
  Dim rsTableName As DAO.Recordset
  Dim rsColumnTypes As New ADODB.Recordset
  Dim rsCommonColumns As New ADODB.Recordset
  Dim strColumnName As String
  
  Set sValueList = New SystemMgr.cStringBuilder
  Set sColumnList = New SystemMgr.cStringBuilder
  fOK = True
  
  With recTabEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While (Not .EOF) And fOK
    
      lngSourceTableID = 0
      lngDestinationTableID = 0
      
      If !copyDataTableID > 0 Then
        lngDestinationTableID = !TableID
        sDestinationTableName = !TableName
        lngSourceTableID = !copyDataTableID
      End If
      
      If (lngSourceTableID > 0) And (lngDestinationTableID > 0) Then
        
        ' Get the source table name.
        sSQL = "SELECT tableName" & _
          " FROM tmpTables" & _
          " WHERE tableID=" & Trim$(Str$(lngSourceTableID))
        Set rsTableName = daoDb.OpenRecordset(sSQL, _
          dbOpenForwardOnly, dbReadOnly)
        If Not (rsTableName.BOF And rsTableName.EOF) Then
          sSourceTableName = rsTableName.Fields("tableName").value
        Else
          fOK = False
        End If
        rsTableName.Close
        
        If fOK Then
          ' Copy the source table into a temporary table.
          sTempCopy = GetTempTableName("Tmp_" & sSourceTableName)
          fOK = Not (sTempCopy = vbNullString)
        End If
          
        If fOK Then

          sSQL = "SELECT * INTO " & sTempCopy & _
            " FROM " & sSourceTableName
          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        
          ' Build list of columns with which to re-populate this table.
          sColumnList.TheString = vbNullString
          sValueList.TheString = vbNullString
          
          ' Get the names of the columns that are common to the source and destination tables.
          sSQL = "SELECT DISTINCT columnName" & _
            " FROM ASRSysColumns " & _
            " WHERE tableID=" & Trim$(Str$(lngDestinationTableID)) & _
            " AND columnName IN " & _
            "   (SELECT columnName" & _
            "     FROM ASRSysColumns" & _
            "     WHERE tableID=" & Trim$(Str$(lngSourceTableID)) & ")"
          rsCommonColumns.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly, adCmdText
          
          With rsCommonColumns
            While Not .EOF
              
              ' Get the datatypes of the source and destination columns.
              strColumnName = .Fields("ColumnName").value
              iSourceColumnDataType = 0
              iDestinationColumnSize = 0
              iDestinationColumnDecimals = 0
              iDestinationColumnDataType = 0
              sSQL = "SELECT tableID, dataType, size, decimals" & _
                " FROM ASRSysColumns" & _
                " WHERE columnName='" & strColumnName & "'" & _
                " AND (tableID=" & Trim$(Str$(lngDestinationTableID)) & _
                " OR tableID=" & Trim$(Str$(lngSourceTableID)) & ")"
              rsColumnTypes.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
                
              With rsColumnTypes
                While Not .EOF
                  If !TableID = lngSourceTableID Then
                    iSourceColumnDataType = !DataType
                  ElseIf !TableID = lngDestinationTableID Then
                    iDestinationColumnDataType = !DataType
                    
                    'TM20060615 - Fault 11085
                    'Specify a size of '14' if a sqlLongVarChar(working pattern) as the size is not
                    'stored in the ASRSysColumns table for this type.
                    If !DataType = SQLDataType.sqlLongVarChar Then
                      iDestinationColumnSize = 14
                    Else
                      iDestinationColumnSize = !Size
                    End If
                    iDestinationColumnDecimals = !Decimals
                  End If
                  .MoveNext
                Wend
                .Close
              End With
              Set rsColumnTypes = Nothing
              
              fOK = (iSourceColumnDataType <> 0) And (iDestinationColumnDataType <> 0)
              
              If fOK Then
                If iDestinationColumnDataType = iSourceColumnDataType Then
                  sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
                  
                  Select Case iDestinationColumnDataType
                    ' Convert character.
                    Case dtVARCHAR, dtLONGVARCHAR
                      sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & _
                        "CONVERT(varchar(" & Trim$(Str$(iDestinationColumnSize)) & ")," & strColumnName & ")"
                      
                    ' Convert numeric.
                    Case dtNUMERIC
                      ' Ensure that we don't try to copy any out of range data into the columns.
                      dblMaxValue = 10 ^ (iDestinationColumnSize - iDestinationColumnDecimals)
                      
                      sSQL = "UPDATE " & sTempCopy & _
                        " SET " & strColumnName & " = 0" & _
                        " WHERE " & strColumnName & " >= " & Trim$(Str$(dblMaxValue)) & _
                        " OR " & strColumnName & " <= -" & Trim$(Str$(dblMaxValue))
                      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

                      sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & _
                        "CONVERT(numeric(" & Trim$(Str$(iDestinationColumnSize)) & "," & Trim$(Str$(iDestinationColumnDecimals)) & "), " & strColumnName & ")"
                    
                    Case Else
                      sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & strColumnName
                      
                  End Select
                
                Else
                  Select Case iDestinationColumnDataType
                    ' Convert data into character if possible.
                    Case dtVARCHAR, dtLONGVARCHAR
                      If (iSourceColumnDataType = dtTIMESTAMP) Or _
                        (iSourceColumnDataType = dtINTEGER) Or _
                        (iSourceColumnDataType = dtNUMERIC) Or _
                        (iSourceColumnDataType = dtBIT) Then
                        sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
                        sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & "CONVERT(varchar(" & Trim$(Str$(iDestinationColumnSize)) & "), " & strColumnName & ")"
                      End If
                                    
                    ' Convert data into integer if possible.
                    Case dtINTEGER
                      If (iSourceColumnDataType = dtNUMERIC) Or _
                        (iSourceColumnDataType = dtBIT) Then
                        sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
                        sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & "CONVERT(int, " & strColumnName & ")"
                      End If
                                  
                    ' Convert data into numeric if possible.
                    Case dtNUMERIC
                      If (iSourceColumnDataType = dtINTEGER) Or _
                        (iSourceColumnDataType = dtBIT) Then
                        sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
                        sValueList.Append IIf(sValueList.Length <> 0, ",", vbNullString) & "CONVERT(numeric(" & Trim$(Str$(iDestinationColumnSize)) & "," & Trim$(Str$(iDestinationColumnDecimals)) & "), " & strColumnName & ")"
                      End If
                                  
                    ' Cannot convert any other datatype into bit, but we need to initialise it.
                    Case dtBIT
                      sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & strColumnName
                      sValueList.Append IIf(sValueList.Length <> 0, ",0", "0")
                  End Select
                End If
              End If
              
              .MoveNext
            Wend
            .Close
          End With
          Set rsCommonColumns = Nothing
                    
          ' Get the names of the logic columns that are only in the destination table (ie. require initialising).
          sSQL = "SELECT DISTINCT columnName" & _
            " FROM ASRSysColumns " & _
            " WHERE tableID=" & Trim$(Str$(lngDestinationTableID)) & _
            " AND dataType=" & Trim$(Str$(dtBIT)) & _
            " AND columnName NOT IN " & _
            "   (SELECT columnName" & _
            "     FROM ASRSysColumns" & _
            "     WHERE tableID=" & Trim$(Str$(lngSourceTableID)) & ")"
          rsCommonColumns.Open sSQL, gADOCon, adOpenDynamic, adLockReadOnly, adCmdText
          
          With rsCommonColumns
            While Not .EOF
              sColumnList.Append IIf(sColumnList.Length <> 0, ",", vbNullString) & !ColumnName
              sValueList.Append IIf(sValueList.Length <> 0, ",0", "0")
              
              .MoveNext
            Wend
            .Close
          End With
          Set rsCommonColumns = Nothing
          
          If (sColumnList.Length <> 0) And (sValueList.Length <> 0) Then
            
            
'MH20010403 The INSERT caused an error which include "IDENTITY INSERT IS OFF".
'           This was strange 'cos we just turned it on in the execute statement above.
'           Anyway, I combined the three execute statements and that seemed to fix
'           the error....... don't ask me!

            ' Populate the destination table with data from the source table.
'            rdoCon.Execute "SET IDENTITY_INSERT " & sDestinationTableName & " ON"
'            sSQL = "INSERT INTO " & sDestinationTableName & " (" & sColumnList & ")" & _
'              " SELECT " & sValueList & " FROM " & sTempCopy
'            rdoCon.Execute sSQL, rdExecDirect
'            rdoCon.Execute "Set IDENTITY_INSERT " & sDestinationTableName & " OFF"
          
            gADOCon.Execute _
                "SET IDENTITY_INSERT " & sDestinationTableName & " ON" & vbNewLine & _
                "INSERT INTO " & sDestinationTableName & " (" & sColumnList.ToString & ")" & _
                " SELECT " & sValueList.ToString & " FROM " & sTempCopy & vbNewLine & _
                "SET IDENTITY_INSERT " & sDestinationTableName & " OFF", , adCmdText + adExecuteNoRecords
          End If
        End If
      End If
      
      .MoveNext
    
    Loop
  End With
  
TidyUpAndExit:

  ' Drop the temporary table.
  If LenB(sTempCopy) <> 0 Then
    sSQL = "IF EXISTS (SELECT Name FROM dbo.sysobjects where id = object_id(N'[dbo].[" & sTempCopy & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" _
          & " DROP TABLE " & sTempCopy
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  End If

  'If Not fOK Then
  '  gobjProgress.Visible = False
  '  MsgBox "Error copying data." & vbCr & vbCr & _
  '    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  'End If
  ' Disassociate object variables.
  Set rsTableName = Nothing
  Set rsCommonColumns = Nothing
  CopyData = fOK
  Exit Function
  
ErrorTrap:
  On Local Error Resume Next
  OutputError "Error copying data"
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function CreateOLEStoredProcedure() As Boolean

  ' Save the new or modified OLE upload stored procedures to the server database.
  On Error GoTo ErrorTrap
  
  Const sSPPrefix = "dbo.spASRUpdateOLEField_"
  Dim fOK As Boolean
  Dim sSQL As String
  Dim strProcedureName As String
  Dim strTableName As String
  
  fOK = True
  
  With recColEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      
      If .Fields("OLEType").value > 1 Then
        strProcedureName = sSPPrefix & .Fields("ColumnID").value
      
        ' Clear existing upload stored procedure
        sSQL = "IF EXISTS" _
          & " (SELECT Name" _
          & "   FROM sysobjects" _
          & "   WHERE id = object_id('" & strProcedureName & "')" _
          & "     AND sysstat & 0xf = 4)" _
          & " DROP PROCEDURE " & strProcedureName
        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
      
        If .Fields("Deleted").value = False Then
          ' Create a nice new fresh stored procedure
          sSQL = "/* ------------------------------------------------ */" & vbNewLine & _
            "/* system stored procedure.                  */" & vbNewLine & _
            "/* Automatically generated by the System manager.   */" & vbNewLine & _
            "/* ------------------------------------------------ */" & vbNewLine & _
            "CREATE PROCEDURE " & strProcedureName & "(" & vbNewLine _
            & "  @piID int," & vbNewLine _
            & "  @pimgUploadFile varbinary(MAX))" & vbNewLine & "AS" & vbNewLine & "BEGIN" & vbNewLine _
            & "  UPDATE " & GetTableName(.Fields("TableID").value) & " SET " _
            & .Fields("ColumnName").value & "= @pimgUploadFile" _
            & " WHERE id =@piID" & vbNewLine _
            & "END"
          gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        End If
      
      End If
      .MoveNext
    Loop
  End With
  
TidyUpAndExit:
  CreateOLEStoredProcedure = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error creating OLE stored procedures"
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function ConfigureModuleSpecifics() As Boolean
  ' Configure module specific objects (eg. stored procedures)
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' If personnel module is enabled
  If Application.PersonnelModule Then
    OutputCurrentProcess2 "Personnel"
    fOK = modPersonnelSpecifics.ConfigurePersonnelSpecifics
  End If
  
  If Application.TrainingBookingModule Then
    OutputCurrentProcess2 "Training Booking"
    fOK = modTrainingBookingSpecifics.ConfigureTrainingBookingSpecifics
  End If

  If fOK And Application.AbsenceModule Then
    OutputCurrentProcess2 "Absence"
    fOK = modAbsenceSpecifics.ConfigureAbsenceSpecifics
  End If

  If fOK Then
    OutputCurrentProcess2 "Maternity"
    fOK = modMaternitySpecifics.ConfigureMaternitySpecifics
  End If

  modWorkflowSpecifics.DropWorkflowObjects
  If Application.WorkflowModule Then
    OutputCurrentProcess2 "Workflow"
    fOK = modWorkflowSpecifics.ConfigureWorkflowSpecifics
  End If

  fOK = modAuditAccess.ConfigureCustomAuditLog
  
  ' NPG20111208
  modMobileSpecifics.DropMobileObjects
  If Application.MobileModule Then
    OutputCurrentProcess2 "Mobile"
    fOK = modMobileSpecifics.ConfigureMobileSpecifics
  End If
  
  ' HRPRO-2303 - reset password and future intranet SP's
  modIntranetSpecifics.DropIntranetObjects
  If Application.SelfServiceIntranetModule Then
    OutputCurrentProcess2 "Intranet"
    fOK = modIntranetSpecifics.ConfigureIntranetSpecifics
  End If
  
  If fOK Then
    OutputCurrentProcess2 "Categories"
    fOK = ConfigureCategories
  End If
  

TidyAndExit:
  ConfigureModuleSpecifics = fOK
Exit Function

ErrorTrap:
  OutputError "Error configuring module specifics"
  fOK = False
  Resume TidyAndExit
  
End Function

Private Function CheckIfRebuildDiaryOrEmail() As Boolean

  Dim strMBText As String
  Dim intMBResponse As VbMsgBoxResult
  Dim rsCount As ADODB.Recordset
  Dim strSQL As String
  Dim blnForceRebuild As Boolean
  Dim lngRecordCount As Long

  On Error GoTo ErrorTrap

  strSQL = "DELETE FROM ASRSysDiaryEvents " & _
           "WHERE ColumnID > 0 AND LinkID NOT IN " & _
           "(SELECT DiaryID FROM ASRSysDiaryLinks)"
  gADOCon.Execute strSQL


  blnForceRebuild = False
  If gfRefreshStoredProcedures Then
    strSQL = "SELECT COUNT(*) FROM ASRSysDiaryEvents WHERE LinkID IS null AND ColumnID > 0"
    Set rsCount = New ADODB.Recordset
    rsCount.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    blnForceRebuild = (rsCount.Fields(0).value > 0)
    rsCount.Close
    Set rsCount = Nothing

    If blnForceRebuild Then
      strSQL = "DELETE FROM ASRSysDiaryEvents WHERE LinkID IS null AND ColumnID > 0"
      gADOCon.Execute strSQL
    End If

  End If


  If Application.ChangedDiaryLink And blnForceRebuild = False Then
    strMBText = "You have made changes which may affect diary events and the diary will have " & _
                "to be rebuilt in order for these changes to take effect." & vbNewLine & vbNewLine & _
                "Would you like to rebuild the diary now?"
    gobjProgress.Visible = False
    Screen.MousePointer = vbDefault
    intMBResponse = MsgBox(strMBText, vbYesNo + vbQuestion, "Diary Rebuild")
  End If
  
  
  gobjProgress.Visible = True
  Screen.MousePointer = vbHourglass
  
  If intMBResponse = vbYes Or blnForceRebuild Then
    OutputCurrentProcess "Rebuilding system diary events"
    DiaryRebuild
    Application.ChangedDiaryLink = False

'    With recTabEdit
'      .Index = "idxTableID"
'      If Not (.BOF And .EOF) Then
'        .MoveFirst
'        lngRecordCount = .RecordCount
'        OutputCurrentProcess2 "", lngRecordCount
'      End If
'      Do While Not .EOF
'
'        'JPD 20040303 Fault 8175
'        'If !New Or !Changed Then
'        If ((!New Or !Changed) And (Not !Deleted)) Or blnForceRebuild Then
'          OutputCurrentProcess2 recTabEdit!TableName
'          strSQL = "EXEC sp_ASRDiaryRebuild_" & CStr(recTabEdit!TableID)
'          gADOCon.Execute strSQL, , adCmdText + adExecuteNoRecords
'          gobjProgress.UpdateProgress2
'        End If
'
'        .MoveNext
'      Loop
'    End With

  End If



  If Application.ChangedEmailLink Then
  
    strMBText = "You have made changes to email link(s) and the email queue will have " & _
                "to be rebuilt in order for these changes to take effect." & vbNewLine & vbNewLine & _
                "The email queue will be rebuilt during the overnight processing however " & _
                "would you like to rebuild the email queue now?"
    gobjProgress.Visible = False
    Screen.MousePointer = vbDefault
    intMBResponse = MsgBox(strMBText, vbYesNo + vbQuestion + vbDefaultButton2, "Email Queue Rebuild")

    gobjProgress.Visible = True
    Screen.MousePointer = vbHourglass
    If intMBResponse = vbYes Then
      Application.ChangedEmailLink = False

      'With recTabEdit
      '  .Index = "idxTableID"
      '  If Not (.BOF And .EOF) Then
      '    .MoveFirst
      '  End If
      '  Do While Not .EOF

          'If !New Or !Changed Or pfRefreshDatabase Then
      '    If !New Or !Changed Then
            OutputCurrentProcess "Rebuilding email queue"
            
            'MH20010305
            'This will probably need to change so that it only rebuilds
            'for the tables which have changed but this will involve changing
            'the way the email rebuild works!
            
            'strSQL = "IF EXISTS (SELECT * FROM sysobjects " & _
                     "WHERE id = object_id('sp_ASREmailRebuild_" & CStr(recTabEdit!TableID) & "') " & _
                     "AND sysstat & 0xf = 4) " & _
                     "BEGIN " & _
                     "EXEC sp_ASRemailRebuild_" & CStr(recTabEdit!TableID) & " " & _
                     "END"
            strSQL = "EXEC dbo.spASREmailRebuild"
            gADOCon.Execute strSQL, , adCmdText + adExecuteNoRecords
      '    End If

      '    .MoveNext
      '  Loop
      'End With

    End If
  
  End If
  
  
  If Application.ChangedOutlookLink Then
  
    strMBText = "You have made changes to calendar link(s) and the Outlook Calendar will have " & _
                "to be rebuilt in order for these changes to take effect." & vbNewLine & vbNewLine & _
                "Would you like to rebuild the Outlook Calendar now?"
    gobjProgress.Visible = False
    Screen.MousePointer = vbDefault
    intMBResponse = MsgBox(strMBText, vbYesNo + vbQuestion + vbDefaultButton2, "Outlook Rebuild")

    gobjProgress.Visible = True
    Screen.MousePointer = vbHourglass
    If intMBResponse = vbYes Then
      Application.ChangedOutlookLink = False

      With recTabEdit
        .Index = "idxTableID"
        If Not (.BOF And .EOF) Then
          .MoveFirst
        End If
        Do While Not .EOF

          'JPD 20040303 Fault 8175
          'If !New Or !Changed Then
          If (!New Or !Changed) And (Not !Deleted) Then
            OutputCurrentProcess "Rebuilding system Outlook events for '" & recTabEdit!TableName & "'"

            strSQL = _
                "IF EXISTS (SELECT Name FROM sysobjects WHERE type = 'P' AND name = 'spASROutlook_" & CStr(recTabEdit!TableID) & "')" & vbNewLine & _
                "BEGIN" & vbNewLine & _
                "  DECLARE @iCurrentID int;" & vbNewLine & _
                "  DECLARE @sSQL nvarchar(MAX);" & vbNewLine & _
                vbNewLine & _
                "  DECLARE curRecords CURSOR FOR" & vbNewLine & _
                "  SELECT id FROM [" & recTabEdit!TableName & "];" & vbNewLine & _
                "  OPEN curRecords;" & vbNewLine & _
                vbNewLine & _
                "  FETCH NEXT FROM curRecords INTO @iCurrentID;" & vbNewLine & _
                "  WHILE @@fetch_status <> -1" & vbNewLine & _
                "  BEGIN" & vbNewLine & _
                "    SET @sSQL = 'EXEC dbo.spASROutlook_" & CStr(recTabEdit!TableID) & " ' + convert(varchar(100), @iCurrentID);" & vbNewLine & _
                "    EXECUTE sp_executeSQL @sSQL;" & vbNewLine & _
                "    FETCH NEXT FROM curRecords INTO @iCurrentID;" & vbNewLine & _
                "  END" & vbNewLine & _
                "  CLOSE curRecords;" & vbNewLine & _
                "  DEALLOCATE curRecords;" & vbNewLine & _
                "END"

            gADOCon.Execute strSQL, , adCmdText + adExecuteNoRecords
          End If

          .MoveNext
        Loop
      End With

    End If
  
  End If
  
  If Application.ChangedWorkflowLink Then
    strMBText = "You have made changes to workflow link(s) and the triggered workflow queue will have " & _
                "to be rebuilt in order for these changes to take effect." & vbNewLine & vbNewLine & _
                "The triggered workflow queue will be rebuilt during the overnight processing however " & _
                "would you like to rebuild the queue now?"
    
    gobjProgress.Visible = False
    Screen.MousePointer = vbDefault
    intMBResponse = MsgBox(strMBText, vbYesNo + vbQuestion + vbDefaultButton2, "Triggered Workflow Queue Rebuild")

    gobjProgress.Visible = True
    Screen.MousePointer = vbHourglass
    
    If intMBResponse = vbYes Then
      Application.ChangedWorkflowLink = False

      OutputCurrentProcess "Rebuilding workflow queue"
      strSQL = "EXEC dbo.spASRWorkflowRebuild"
      gADOCon.Execute strSQL, , adCmdText + adExecuteNoRecords
    End If
  End If
  
  CheckIfRebuildDiaryOrEmail = True

Exit Function

ErrorTrap:
  OutputError "Error Rebuilding Diary and/or Email Queue"
  'Still return true to disable the save button
  '(because the save process completed ok!)
  CheckIfRebuildDiaryOrEmail = True

End Function

Private Function UpdateLockCheck() As Boolean
  
  Dim blnCurrentlyLocked As Boolean
  Dim strLockDetails As String
  
  Dim rsTemp As ADODB.Recordset
  
  Set rsTemp = New ADODB.Recordset
  rsTemp.Open "sp_ASRLockCheck", gADOCon, adOpenDynamic, adLockReadOnly, adCmdStoredProc
  
  If Not (rsTemp.BOF And rsTemp.EOF) Then
    'Ignore users own manual lock
    If LCase(gsUserName) = LCase(rsTemp!userName) And rsTemp!Priority = lckManual Then
      rsTemp.MoveNext
    End If
    
    If Not (rsTemp.BOF And rsTemp.EOF) Then
  
      'If not locked by current app then can we get read only access...
      strLockDetails = "User :  " & rsTemp!userName & vbNewLine & _
                       "Date/Time :  " & rsTemp!Lock_Time & vbNewLine & _
                       "Machine :  " & rsTemp!HostName & vbNewLine & _
                       "Type :  " & rsTemp!Description
  
      Screen.MousePointer = vbDefault
      
      'Database is locked
      MsgBox "Unable to update '" & gsDatabaseName & "' as the database has been locked." & _
             vbNewLine & vbNewLine & strLockDetails, vbExclamation, Application.Name
      
      Application.AccessMode = accNone
      Screen.MousePointer = vbHourglass

    End If
  
  End If
    
  rsTemp.Close
  Set rsTemp = Nothing
End Function

Private Function RunRecordSaveOptimiser() As Boolean

  Dim strSQL As String
  Dim bOK As Boolean

  On Error GoTo ErrorTrap

  bOK = True
  strSQL = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('spadmin_optimiserecordsave')" & _
    "     AND sysstat & 0xf = 4)" & _
    " EXECUTE sp_executeSQL spadmin_optimiserecordsave;"
    
 ' gADOCon.Execute strSQL, , adExecuteNoRecords

  RunRecordSaveOptimiser = bOK
Exit Function

ErrorTrap:
  OutputError "Error Removing overnight process"
  RunRecordSaveOptimiser = False

End Function


Private Function SaveMobileNavigation() As Boolean

  Dim sValues As String
  Dim sInsert As String
  Dim sSQL As String
  Dim bOK As Boolean
  Dim objTable As DAO.TableDef
  Dim objField As DAO.Field
  Dim objRecordset As DAO.Recordset

  On Error GoTo ErrorTrap

  bOK = True
 
  bOK = bOK And SquirtDAOTableToSQL("tmpmobileformlayout", "tbsys_mobileformlayout")
  bOK = bOK And SquirtDAOTableToSQL("tmpmobileformelements", "tbsys_mobileformelements")
  bOK = bOK And SquirtDAOTableToSQL("tmpmobilegroupworkflows", "tbsys_mobilegroupworkflows")
  
TidyUpAndExit:
  SaveMobileNavigation = bOK
  Exit Function

ErrorTrap:
  OutputError "Error Saving Mobile Navigation"
  bOK = False
  Resume TidyUpAndExit


End Function

' An incredibly inefficient way of hoofing data from the Access DB to SQL. As soon as we can move to NHibernate the better!
Public Function SquirtDAOTableToSQL(ByRef strAccessTable As String, ByRef strSQLTable As String) As Boolean

  Dim sValues As String
  Dim sInsert As String
  Dim sSQL As String
  Dim bOK As Boolean
  Dim objTable As DAO.TableDef
  Dim objField As DAO.Field
  Dim objRecordset As DAO.Recordset

  On Error GoTo ErrorTrap

  bOK = True

  ' Get rid of existing data
  gADOCon.Execute "DELETE FROM " & strSQLTable, , adCmdText + adExecuteNoRecords

  ' Build insert string
  sInsert = ""
  Set objTable = daoDb.TableDefs(strAccessTable)
  For Each objField In objTable.Fields
    sInsert = sInsert & ", " & objField.Name
  Next objField
  sInsert = Mid(sInsert, 2)
  sInsert = "INSERT [" & strSQLTable & "] (" & sInsert & ") VALUES ("
  
  ' Build each of the update commands
  Set objRecordset = daoDb.OpenRecordset("SELECT * FROM " & strAccessTable)
  If objRecordset.EOF And objRecordset.BOF Then
    SquirtDAOTableToSQL = True
    Exit Function
  End If
  
  objRecordset.MoveFirst
  Do While Not objRecordset.EOF
  
    sValues = ""
    For Each objField In objRecordset.Fields
    
      Select Case objField.Type
        Case 1 ' Boolean
          sValues = sValues & ", " & CStr(IIf(IsNull(objField.value), "NULL", IIf(objField.value, "1", "0")))
         
        Case 2, 7 ' Number
          sValues = sValues & ", " & CStr(IIf(IsNull(objField.value), "NULL", objField.value))
               
        Case 10, 12 ' Text
          If IsNull(objField.value) Then
            sValues = sValues & ", NULL"
          Else
            sValues = sValues & ", '" & Replace(objField.value, "'", "''") & "'"
          End If
          
        Case 4 ' Date
          sValues = sValues & ", " & CStr(IIf(IsNull(objField.value), "NULL", objField.value))
          
        Case Else
          sValues = sValues & ", " & CStr(IIf(IsNull(objField.value), "NULL", objField.value))
      
      End Select
    Next objField
    sValues = Mid(sValues, 2)
    
    sSQL = sInsert & sValues & ")"

    ' Insert into SQL
    gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

    objRecordset.MoveNext
  Loop

TidyUpAndExit:
  SquirtDAOTableToSQL = bOK
  Exit Function

ErrorTrap:
  OutputError "Error Saving Mobile Navigation"
  bOK = False
  Resume TidyUpAndExit

End Function

Private Function ConfigureCategories() As Boolean

  Dim bOK As Boolean
  Dim sSQL As String
  Dim sCategoryTableName As String
  Dim sCategoryColumnName As String
      
  ' Drop existing stuff
  bOK = DropProcedure("spsys_getobjectcategories")
  bOK = DropFunction("udfsys_getcategory")
  bOK = DropView("ASRSysCategories")

  ' Module defined columns
  sCategoryTableName = GetTableName(GetModuleSetting(gsMODULEKEY_CATEGORY, gsPARAMETERKEY_CATEGORYTABLE, 0))
  sCategoryColumnName = GetColumnName(GetModuleSetting(gsMODULEKEY_CATEGORY, gsPARAMETERKEY_CATEGORYNAMECOLUMN, 0), True)
   
  ' Funky new procedure
  sSQL = "/* ---------------------------------------------------- */" & vbNewLine & _
            "/* Categories module stored procedure.               */" & vbNewLine & _
            "/* Automatically generated by the System manager.    */" & vbNewLine & _
            "/* ---------------------------------------------------- */" & vbNewLine & _
            "CREATE PROCEDURE dbo.[spsys_getobjectcategories](@utilityType as integer, @UtilityID as integer, @tableID as integer)" & vbNewLine & _
            "AS" & vbNewLine & _
            "BEGIN" & vbNewLine & _
            "    SET NOCOUNT ON;" & vbNewLine & vbNewLine
    
  If Len(sCategoryTableName) > 0 And Len(sCategoryColumnName) > 0 Then
    sSQL = sSQL & "    SELECT c.ID, c.[" & sCategoryColumnName & "] AS [category_name]" & vbNewLine & _
              "        , CASE ISNULL(s.categoryid,0) WHEN 0 THEN 0 ELSE 1 END   AS [selected]" & vbNewLine & _
              "        FROM dbo.[" & sCategoryTableName & "] c" & vbNewLine & _
              "            LEFT JOIN tbsys_objectcategories s ON s.CategoryID = c.ID AND s.objecttype = @utilityType AND s.objectid = @UtilityID" & vbNewLine & _
              "        ORDER BY c.[" & sCategoryColumnName & "];" & vbNewLine
  Else
    sSQL = sSQL & "    SELECT 0 AS ID, '' AS [category_name], 0 AS [selected] WHERE 1=2;" & vbNewLine
  End If
    
  sSQL = sSQL & "END"
    
  gADOCon.Execute sSQL, , adExecuteNoRecords
     

  sSQL = "/* ---------------------------------------------------- */" & vbNewLine & _
            "/* Categories module function.               */" & vbNewLine & _
            "/* Automatically generated by the System manager.    */" & vbNewLine & _
            "/* ---------------------------------------------------- */" & vbNewLine & _
            "CREATE FUNCTION dbo.[udfsys_getcategory](@categoryID as integer)" & vbNewLine & _
            "RETURNS nvarchar(MAX)" & vbNewLine & _
            "AS" & vbNewLine & _
            "BEGIN" & vbNewLine & _
            "    DECLARE @result nvarchar(MAX);" & vbNewLine & vbNewLine & _
            "    SELECT @result = [" & sCategoryColumnName & "] FROM dbo.[" & sCategoryTableName & "] WHERE ID = @categoryID" & vbNewLine & _
            "    RETURN @result" & vbNewLine & _
            "END"
  gADOCon.Execute sSQL, , adExecuteNoRecords
     
     
  sSQL = "/* ---------------------------------------------------- */" & vbNewLine & _
            "/* Categories module view.                           */" & vbNewLine & _
            "/* Automatically generated by the System manager.    */" & vbNewLine & _
            "/* ---------------------------------------------------- */" & vbNewLine & _
            "CREATE VIEW dbo.[ASRSysCategories]" & vbNewLine & _
            "AS" & vbNewLine & _
            "    SELECT ID, [" & sCategoryColumnName & "] AS Category_Name FROM dbo.[tbuser_" & sCategoryTableName & "]"
  gADOCon.Execute sSQL, , adExecuteNoRecords
     
     
  ConfigureCategories = bOK

End Function

Private Function ApplyPostSaveProcessing() As Boolean

  On Error GoTo ErrorTrap

  Dim cmdPostProcess As New ADODB.Command
  Dim bOK As Boolean

  bOK = True
  
  With cmdPostProcess
    .CommandText = "spASRPostSystemSave"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    .Execute
  End With

  Set cmdPostProcess = Nothing

TidyUpAndExit:
  ApplyPostSaveProcessing = bOK
  Exit Function

ErrorTrap:
  OutputError "Error Applying Post Save Processing"
  bOK = False
  GoTo TidyUpAndExit

End Function
