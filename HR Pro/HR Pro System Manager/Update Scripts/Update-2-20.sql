
/* ----------------------------------------------------- */
/* Update the database from version 2.19 to version 2.20 */
/* ----------------------------------------------------- */

DECLARE @iRecCount integer,
	@iType integer,
	@iLength integer,
	@sDBVersion varchar(10),
	@sCommand nvarchar(500),
	@sParam	nvarchar(500),
	@sName sysname,
	@ptrval binary(16),
	@DBName varchar(255),
	@Command varchar(8000),
	@GroupName varchar(8000),
	@NVarCommand nvarchar(4000),
	@sColumnDataType varchar(8000),
	@iDateFormat varchar(255),
	@sSQLVersion nvarchar(20),
	@iTemp integer,
	@sTemp varchar(8000)

DECLARE @sGroup sysname
DECLARE @sObject sysname
DECLARE @sObjectType char(2)
DECLARE @sSQL varchar(8000)


/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not version 2.17 or 2.18. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '2.19') and (@sDBVersion <> '2.20')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END



/* ------------------------------------------------------------- */
PRINT 'Step 1 of 2 - Amending Outlook Calendar Procedure'

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASROutlookBatch]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASROutlookBatch]

		EXEC('CREATE PROCEDURE dbo.spASROutlookBatch
		AS
		BEGIN
		
		DECLARE @sSQL nvarchar(4000)
		DECLARE @sParamDefinition nvarchar(4000)
		DECLARE @hResult int
		DECLARE @objectToken integer
		DECLARE @temp int
		DECLARE @CharValue varchar(8000)
		DECLARE @DateValue datetime
		
		DECLARE @StoreID varchar(8000)
		DECLARE @EntryID varchar(8000)
		DECLARE @StoreIDTemp varchar(8000)
		DECLARE @EntryIDTemp varchar(8000)
		DECLARE @Start int
		DECLARE @Index int
	
		DECLARE @LinkID int
		DECLARE @TableID int
		DECLARE @RecordID int
		DECLARE @Refresh bit
		DECLARE @Deleted bit
		DECLARE @Title varchar(8000)
		DECLARE @BusyStatus int
		DECLARE @StartDateColumnID int
		DECLARE @EndDateColumnID int
		DECLARE @TimeRange int
		DECLARE @FixedStartTime varchar(8000)
		DECLARE @FixedEndTime varchar(8000)
		DECLARE @StartTimeColumnID int
		DECLARE @EndTimeColumnID int
		DECLARE @SubjectExprID int
		DECLARE @Content varchar(8000)
		DECLARE @Reminder bit
		DECLARE @ReminderOffset int
		DECLARE @ReminderPeriod int
		DECLARE @FolderID int
		DECLARE @FolderType int
		DECLARE @FolderPath varchar(8000)
		DECLARE @FolderExprID int
		DECLARE @RecordDescExprID int
		DECLARE @Subject varchar(8000)
		
		DECLARE @ErrorMessage varchar(8000)
		
		DECLARE @Folder varchar(8000)
		DECLARE @AllDayEvent bit
		DECLARE @StartDate datetime
		DECLARE @EndDate datetime
		DECLARE @StartTime varchar(8000)
		DECLARE @EndTime varchar(8000)
		
		DECLARE @Heading varchar(8000)
		DECLARE @TableName varchar(8000)
		DECLARE @ColumnName varchar(8000)
		DECLARE @DataType int
	
		DECLARE @DateFormat varchar(8000)
	
		DECLARE CursorEvents CURSOR FOR 
		SELECT ASRSysOutlookEvents.LinkID,
			ASRSysOutlookEvents.FolderID,
			ASRSysOutlookEvents.TableID,
			ASRSysOutlookEvents.RecordID,
			ASRSysOutlookEvents.Refresh,
			ASRSysOutlookEvents.Deleted,
			ASRSysOutlookEvents.StoreID,
			ASRSysOutlookEvents.EntryID,
			ASRSysOutlookLinks.Title,
			ASRSysOutlookLinks.BusyStatus,
			ASRSysOutlookLinks.StartDate,
			ASRSysOutlookLinks.EndDate,
			ASRSysOutlookLinks.TimeRange,
			ASRSysOutlookLinks.FixedStartTime,
			ASRSysOutlookLinks.FixedEndTime,
			ASRSysOutlookLinks.ColumnStartTime,
			ASRSysOutlookLinks.ColumnEndTime,
			ASRSysOutlookLinks.Subject,
			isnull(ASRSysOutlookLinks.Content,''''),
			ASRSysOutlookLinks.Reminder,
			ASRSysOutlookLinks.ReminderOffset,
			ASRSysOutlookLinks.ReminderPeriod,
			ASRSysOutlookFolders.FolderType,
			ASRSysOutlookFolders.FixedPath,
			ASRSysOutlookFolders.ExprID,
			ASRSysTables.RecordDescExprID
		FROM ASRSysOutlookEvents
		LEFT OUTER JOIN ASRSysOutlookLinks
			ON ASRSysOutlookEvents.LinkID = ASRSysOutlookLinks.LinkID
		LEFT OUTER JOIN ASRSysOutlookFolders
			ON ASRSysOutlookEvents.FolderID = ASRSysOutlookFolders.FolderID
		LEFT OUTER JOIN ASRSysTables
			ON ASRSysOutlookEvents.TableID = ASRSysTables.TableID
		WHERE ASRSysOutlookEvents.Refresh = 1
			OR ASRSysOutlookEvents.Deleted = 1
		
		
		OPEN CursorEvents
		FETCH NEXT FROM CursorEvents
		INTO	@LinkID, @FolderID, @TableID, @RecordID, @Refresh, @Deleted, @StoreID, @EntryID,
			@Title, @BusyStatus, @StartDateColumnID, @EndDateColumnID, @TimeRange,
			@FixedStartTime, @FixedEndTime, @StartTimeColumnID, @EndTimeColumnID,
			@SubjectExprID, @Content, @Reminder, @ReminderOffset, @ReminderPeriod,
			@FolderType, @FolderPath, @FolderExprID, @RecordDescExprID
	
		IF @@FETCH_STATUS = 0
		BEGIN
	
			SELECT @DateFormat = SettingValue FROM ASRSysSystemSettings WHERE [Section] = ''email'' AND [SettingKey] = ''date format''
	
			EXEC @hResult = sp_OACreate ''vbpHRProServer.clsOutlookCalendar'', @objectToken OUTPUT
			IF @hResult = 0
				EXEC @hResult = sp_OAMethod @objectToken, ''Logon'', @temp OUTPUT, '''', ''''
		
			WHILE @@FETCH_STATUS = 0
			BEGIN
			
				EXEC @hResult = sp_OAMethod @objectToken, ''ResetStoreAndEntry'', @temp OUTPUT
		
				IF @hResult = 0 AND isnull(@StoreID,'''') <> '''' AND isnull(@EntryID,'''') <> ''''
				BEGIN
					SET @Index = 0
					WHILE @Index < 9
					BEGIN
						SET @Start = (@Index * 255) + 1
						IF @Start < LEN(@StoreID)
						BEGIN
							SET @StoreIDTemp = substring(@StoreID,@Start,255)
							EXEC @hResult = sp_OAMethod @objectToken, ''AddToStoreID'', @temp OUTPUT, @StoreIDTemp
						END
						IF @Start < LEN(@EntryID)
						BEGIN
							SET @EntryIDTemp = substring(@EntryID,@Start,255)
							EXEC @hResult = sp_OAMethod @objectToken, ''AddToEntryID'', @temp OUTPUT, @EntryIDTemp
						END
						SET @Index = @Index + 1
					END
		
					EXEC @hResult = sp_OAMethod @objectToken, ''DeleteEntry'', @temp OUTPUT
					EXEC @hResult = sp_OAGetProperty @objectToken, ''ErrorMessage'', @ErrorMessage OUTPUT
				END
		
			
				SET @ErrorMessage = ''''
				SET @StoreID = ''''
				SET @EntryID = ''''
			
				IF @Deleted = 1
				BEGIN
					DELETE FROM ASRSysOutlookEvents WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID
				END
				ELSE
				BEGIN
			
					SELECT @sSQL = ''SELECT @StartDate=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
					FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
					WHERE ColumnID = @StartDateColumnID
					SET @sParamDefinition = N''@StartDate datetime OUTPUT''
					EXEC sp_executesql @sSQL,  @sParamDefinition, @StartDate OUTPUT
			
					SET @EndDate = Null
					IF @EndDateColumnID > 0
					BEGIN
						SELECT @sSQL = ''SELECT @EndDate=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
						FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
						WHERE ColumnID = @EndDateColumnID
						SET @sParamDefinition = N''@EndDate datetime OUTPUT''
						EXEC sp_executesql @sSQL,  @sParamDefinition, @EndDate OUTPUT
						IF rtrim(@EndDate) = '''' SET @EndDate = null
					END
			
					IF @TimeRange = 0
					BEGIN
						SET @AllDayEvent = 1
						SET @StartTime = ''''
						SET @EndTime = ''''
					END
					IF @TimeRange = 1
					BEGIN
						SET @AllDayEvent = 0
						SET @StartTime = @FixedStartTime
						SET @EndTime = @FixedEndTime
					END
					IF @TimeRange = 2
					BEGIN
						SET @AllDayEvent = 0
			
						SELECT @sSQL = ''SELECT @StartTime=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
						FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
						WHERE ColumnID = @StartTimeColumnID
						SET @sParamDefinition = N''@StartTime varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL,  @sParamDefinition, @StartTime OUTPUT
			
						SELECT @sSQL = ''SELECT @EndTime=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
						FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
						WHERE ColumnID = @EndTimeColumnID
						SET @sParamDefinition = N''@EndTime varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL,  @sParamDefinition, @EndTime OUTPUT
			
						IF UPPER(@StartTime) = ''AM''
							SELECT @StartTime = SettingValue FROM ASRSysSystemSettings
							WHERE [Section] = ''outlook'' and [Settingkey] = ''amstarttime''
						IF UPPER(@StartTime) = ''PM''
							SELECT @StartTime = SettingValue FROM ASRSysSystemSettings
							WHERE [Section] = ''outlook'' and [Settingkey] = ''pmstarttime''
						IF UPPER(@EndTime) = ''AM''
							SELECT @EndTime = SettingValue FROM ASRSysSystemSettings
							WHERE [Section] = ''outlook'' and [Settingkey] = ''amendtime''
						IF UPPER(@EndTime) = ''PM''
							SELECT @EndTime = SettingValue FROM ASRSysSystemSettings
							WHERE [Section] = ''outlook'' and [Settingkey] = ''pmendtime''
					END
			
			
					SET @Subject = ''''
					IF @SubjectExprID > 0
					BEGIN
						SET @sSQL = ''DECLARE @hResult int
							IF EXISTS(SELECT * FROM sysobjects WHERE type = ''''P'''' AND name = ''''sp_ASRExpr_''+convert(nvarchar(4000),@SubjectExprID)+'''''')
						             BEGIN
						                EXEC @hResult = sp_ASRExpr_''+convert(nvarchar(4000),@SubjectExprID)+'' @Subject OUTPUT, ''+convert(nvarchar(4000),@RecordID)+''
						                IF @hResult <> 0 SET @Subject = ''''''''
						                SET @Subject = CONVERT(varchar(255), @Subject)
							     END
							     ELSE SET @Subject = ''''''''''
						SET @sParamDefinition = N''@Subject varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL,  @sParamDefinition, @Subject OUTPUT
					END
					ELSE
					BEGIN
						IF @RecordDescExprID > 0
						BEGIN
							SET @sSQL = ''DECLARE @hResult int
								IF EXISTS(SELECT * FROM sysobjects WHERE type = ''''P'''' AND name = ''''sp_ASRExpr_''+convert(nvarchar(4000),@RecordDescExprID)+'''''')
							             BEGIN
							                EXEC @hResult = sp_ASRExpr_''+convert(nvarchar(4000),@RecordDescExprID)+'' @Subject OUTPUT, ''+convert(nvarchar(4000),@RecordID)+''
							                IF @hResult <> 0 SET @Subject = ''''''''
							                SET @Subject = CONVERT(varchar(255), @Subject)
								     END
								     ELSE SET @Subject = ''''''''''
							SET @sParamDefinition = N''@Subject varchar(8000) OUTPUT''
							EXEC sp_executesql @sSQL,  @sParamDefinition, @Subject OUTPUT
							IF @Subject <> ''''
								SET @Subject = '': ''+@Subject
						END
						SET @Subject = @Title+@Subject
					END
			
			
					SET @Folder = @FolderPath
					IF @FolderType > 0
					BEGIN
						SET @sSQL = ''DECLARE @hResult int
							IF EXISTS(SELECT * FROM sysobjects WHERE type = ''''P'''' AND name = ''''sp_ASRExpr_''+convert(nvarchar(4000),@FolderExprID)+'''''')
						             BEGIN
						                EXEC @hResult = sp_ASRExpr_''+convert(nvarchar(4000),@FolderExprID)+'' @Folder OUTPUT, ''+convert(nvarchar(4000),@RecordID)+''
						                IF @hResult <> 0 SET @Folder = ''''''''
							     END
							     ELSE SET @Folder = ''''''''''
						SET @sParamDefinition = N''@Folder varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL,  @sParamDefinition, @Folder OUTPUT
					END
			
			
					DECLARE CursorColumns CURSOR FOR 
					SELECT isnull(ASRSysOutlookLinksColumns.Heading,''''),
						ASRSysTables.TableName,
						ASRSysColumns.ColumnName,
						ASRSysColumns.DataType
					FROM ASRSysOutlookLinksColumns
					JOIN ASRSysColumns
						ON ASRSysColumns.ColumnID = ASRSysOutlookLinksColumns.ColumnID
					JOIN ASRSysTables
						ON ASRSysColumns.TableID = ASRSysTables.TableID
					WHERE LinkID = @LinkID
					ORDER BY [Sequence] DESC
			
					SET @Content = char(13) + @Content
			
					OPEN CursorColumns
					FETCH NEXT FROM CursorColumns
					INTO	@Heading, @TableName, @ColumnName, @DataType
			
					WHILE @@FETCH_STATUS = 0
					BEGIN
	
						IF @Heading <> '''' SET @Heading = @Heading+'': ''
	
						IF @DataType = 12
							SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+isnull([''+@ColumnName+''],'''''''') FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
						IF @DataType = 11
							SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+case when [''+@ColumnName+''] is null then ''''<Empty>'''' else convert(varchar(8000),[''+@ColumnName+''],''+@DateFormat+'') end FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
						IF @DataType = -7
							SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+case when [''+@ColumnName+''] = 1 then ''''Y'''' else ''''N'''' end FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
						IF @DataType <> 11 AND @DataType <> 12 AND @DataType <> -7
							SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+convert(varchar(8000),isnull([''+@ColumnName+''],'''''''')) FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
	
						SET @sParamDefinition = N''@CharValue varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL,  @sParamDefinition, @CharValue OUTPUT
			
						IF @CharValue IS Null SET @CharValue = ''''
						SET @Content = @CharValue + char(13) + @Content
			
						FETCH NEXT FROM CursorColumns
						INTO	@Heading, @TableName, @ColumnName, @DataType
					END
			
					CLOSE CursorColumns
					DEALLOCATE CursorColumns
		
		
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''Reminder'', @Reminder
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''ReminderOffset'', @ReminderOffset
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''ReminderPeriod'', @ReminderPeriod
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''AllDayEvent'', @AllDayEvent
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''StartDate'', @StartDate
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''StartTime'', @StartTime
					IF @hResult = 0 AND NOT @EndDate IS NULL
						EXEC @hResult = sp_OASetProperty @objectToken, ''EndDate'', @EndDate
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''EndTime'', @EndTime
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''Subject'', @Subject
					IF @hResult = 0 AND NOT @Content IS NULL
						EXEC @hResult = sp_OASetProperty @objectToken, ''Content'', @Content
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''BusyStatus'', @BusyStatus
					IF @hResult = 0
						EXEC @hResult = sp_OASetProperty @objectToken, ''Folder'', @Folder
					IF @hResult = 0
						EXEC @hResult = sp_OAMethod @objectToken, ''CreateEntry'', @temp OUTPUT
						IF @hResult = 0
							BEGIN
								--Need to return the STOREID in chunks as SQL7 has a problem returning long strings
								SET @StoreID = ''''
								SET @Index = 0
								WHILE @Index < 9
								BEGIN
									EXEC @hResult = sp_OAGetProperty @objectToken, ''StoreID'', @StoreIDTemp OUTPUT, @Index
									IF @StoreIDTemp <> '''' SET @StoreID = @StoreID + @StoreIDTemp
									SET @Index = @Index + 1
								END
							END
						IF @hResult = 0
							BEGIN
								--Need to return the STOREID in chunks as SQL7 has a problem returning long strings
								SET @EntryID = ''''
								SET @Index = 0
								WHILE @Index < 9
								BEGIN
									EXEC @hResult = sp_OAGetProperty @objectToken, ''EntryID'', @EntryIDTemp OUTPUT, @Index
									IF @EntryIDTemp <> '''' SET @EntryID = @EntryID + @EntryIDTemp
									SET @Index = @Index + 1
								END
							END
					IF @hResult = 0
						EXEC @hResult = sp_OAGetProperty @objectToken, ''ErrorMessage'', @ErrorMessage OUTPUT
					IF @hResult = 0
						EXEC @hResult = sp_OAGetProperty @objectToken, ''StartDate'', @StartDate OUTPUT
					IF @hResult = 0
						EXEC @hResult = sp_OAGetProperty @objectToken, ''EndDate'', @EndDate OUTPUT
		
					IF @hResult <> 0 
					BEGIN
						EXEC sp_OAGetErrorInfo @objectToken, '''', @ErrorMessage OUTPUT
						SET @ErrorMessage = ''Outlook Error: ''+rtrim(ltrim(@ErrorMessage))
						RAISERROR (@ErrorMessage,3,1)
					END
		
					UPDATE ASRSysOutlookEvents
					SET ErrorMessage = @ErrorMessage, StoreID = @StoreID, EntryID = @EntryID, Refresh = 0, StartDate = @StartDate, Subject = @Subject, Folder = @Folder, EndDate = @EndDate, RefreshDate = getdate()
					WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID
		
				END
		
		
				FETCH NEXT FROM CursorEvents
				INTO	@LinkID, @FolderID, @TableID, @RecordID, @Refresh, @Deleted, @StoreID, @EntryID,
					@Title, @BusyStatus, @StartDateColumnID, @EndDateColumnID, @TimeRange,
					@FixedStartTime, @FixedEndTime, @StartTimeColumnID, @EndTimeColumnID,
					@SubjectExprID, @Content, @Reminder, @ReminderOffset, @ReminderPeriod,
					@FolderType, @FolderPath, @FolderExprID, @RecordDescExprID
			END
		
			EXEC @hResult = sp_OAMethod @objectToken, ''Quit'', @temp OUTPUT
			EXEC sp_OADestroy @objectToken
	
		END
	
		CLOSE CursorEvents
		DEALLOCATE CursorEvents
		
		
		END')


/* ------------------------------------------------------------- */
PRINT 'Step X of X - Adding new column to Self-service Intranet Views'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysSSIViews')
		and name = 'pageTitle'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIViews ADD
			                       [pageTitle] [varchar] (200) NULL'
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysSSIViews
				SET pageTitle = ''Select the required '' 
					+ (SELECT replace(ASRSysViews.viewName, ''_'', '' '')
						FROM ASRSysViews
						WHERE ASRSysViews.viewID = ASRSysSSIViews.viewID)
					+ '' record'''
			EXEC sp_executesql @NVarCommand

			SELECT @NVarCommand = 'UPDATE ASRSysSSIViews
				SET pageTitle = ''''
				WHERE pageTitle IS NULL'
			EXEC sp_executesql @NVarCommand


		END

/* ------------------------------------------------------------- */
PRINT 'Step X of X - Adding new columns to Workflow Instance Steps'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
		and name = 'ActivationDateTime'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD
			                       [ActivationDateTime] [datetime] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
		and name = 'Message'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD
			                       [Message] [varchar] (8000) NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceSteps')
		and name = 'CompletionDateTime'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceSteps ADD
			                       [CompletionDateTime] [datetime] NULL'
			EXEC sp_executesql @NVarCommand
		END

/* ------------------------------------------------------------- */
PRINT 'Step X of X - Adding new columns to Workflow Instances'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstances')
		and name = 'InitiationDateTime'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstances ADD
			                       [InitiationDateTime] [datetime] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstances')
		and name = 'CompletionDateTime'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstances ADD
			                       [CompletionDateTime] [datetime] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstances')
		and name = 'UserName'
		
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstances ADD
			                       [UserName] [varchar] (256) NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysWorkflowInstances')
		and name = 'Status'
		
		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstances ADD
			                       [Status] [int] NULL'
			EXEC sp_executesql @NVarCommand
		END

/* ------------------------------------------------------------- */
PRINT 'Step X of X - Modifying column in Workflow Instance Values'
/* NB. We assume at this point that no workflows have been instantiated. 
Hence, we alter the ID column, but do not try to retain values already in it. */
		Set @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues DROP COLUMN ID'
		EXEC sp_executesql @NVarCommand

		Set @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD
			                       [ID] [int] IDENTITY (1, 1) NOT NULL'
		EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step X of X - Creating Workflow Procedures'

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetActiveWorkflowStoredDataSteps]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRGetActiveWorkflowStoredDataSteps]

		EXEC('CREATE PROCEDURE dbo.spASRGetActiveWorkflowStoredDataSteps
			AS
			BEGIN
				/* Return a recordset of the workflow StoredData steps that need to be actioned by the Workflow service. */
				SELECT S.instanceID AS [instanceID],
					E.ID AS [elementID],
					S.ID AS [stepID]
				FROM ASRSysWorkflowInstanceSteps S
				INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
				WHERE S.status = 1
					AND E.type = 5 -- 5 = Stored Data
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetDecisionSucceedingWorkflowElements]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRGetDecisionSucceedingWorkflowElements]

		EXEC('CREATE PROCEDURE dbo.spASRGetDecisionSucceedingWorkflowElements
			(
				@piElementID		integer,
				@piValue		integer,
				@succeedingElements	cursor varying output
			)
			AS
			BEGIN
				CREATE TABLE #succeedingElements (elementID integer)

				/* Return the cursor of succeeding elements. */
				SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
					SELECT elementID 
					FROM #succeedingElements
				OPEN @succeedingElements
			
				DROP TABLE #succeedingElements
			END')

		EXEC('ALTER PROCEDURE dbo.spASRGetDecisionSucceedingWorkflowElements
			(
				@piElementID		integer,
				@piValue		integer,
				@succeedingElements	cursor varying output
			)
			AS
			BEGIN
				/* Return the IDs of the workflow elements that succeed the given element.
				This ignores connection elements.
				NB. This does work for elements with multiple outbound flows. */
				DECLARE
					@iConnectorPairID	integer,
					@superCursor		cursor,
					@iTemp		integer
				
				CREATE TABLE #succeedingElements (elementID integer)
			
				/* Get the non-connector elements. */
				INSERT INTO #succeedingElements
				SELECT L.endElementID
				FROM ASRSysWorkflowLinks L
				INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
				WHERE L.startElementID = @piElementID
					AND L.startOutboundFlowCode = @piValue
					AND E.type <> 8 -- 8 = Connector 1
			
				DECLARE succeedingConnectorsCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.connectionPairID
				FROM ASRSysWorkflowLinks L
				INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
				WHERE L.startElementID = @piElementID
					AND L.startOutboundFlowCode = @piValue
					AND E.type = 8 -- 8 = Connector 1
			
				OPEN succeedingConnectorsCursor
				FETCH NEXT FROM succeedingConnectorsCursor INTO @iConnectorPairID
				WHILE (@@fetch_status = 0)
				BEGIN
					EXEC spASRGetDecisionSucceedingWorkflowElements @iConnectorPairID, @piValue, @superCursor OUTPUT
			
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
						
						FETCH NEXT FROM @superCursor INTO @iTemp 
					END
					CLOSE @superCursor
					DEALLOCATE @superCursor
			
					FETCH NEXT FROM succeedingConnectorsCursor INTO @iConnectorPairID
				END
				CLOSE succeedingConnectorsCursor
				DEALLOCATE succeedingConnectorsCursor
			
				/* Return the cursor of succeeding elements. */
				SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
					SELECT elementID 
					FROM #succeedingElements
				OPEN @succeedingElements
			
				DROP TABLE #succeedingElements
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetPrecedingWorkflowElements]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRGetPrecedingWorkflowElements]

		EXEC('CREATE PROCEDURE dbo.spASRGetPrecedingWorkflowElements
			(
				@piElementID		integer,
				@precedingElements	cursor varying output
			)
			AS
			BEGIN
				CREATE TABLE #precedingElements (elementID integer)

				/* Return the cursor of preceding elements. */
				SET @precedingElements = CURSOR FORWARD_ONLY STATIC FOR
					SELECT elementID 
					FROM #precedingElements
				OPEN @precedingElements
			
				DROP TABLE #precedingElements
			END')

		EXEC('ALTER PROCEDURE dbo.spASRGetPrecedingWorkflowElements
			(
				@piElementID		integer,
				@precedingElements	cursor varying output
			)
			AS
			BEGIN
				/* Return the IDs of the workflow elements that precede the given element.
				This bypasses connection elements.
				NB. This does work for elements with multiple inbound flows. */
				DECLARE
					@iConnectorPairID	integer,
					@superCursor		cursor,
					@iTemp		integer
			
				CREATE TABLE #precedingElements (elementID integer)
			
				/* Get the non-connector elements. */
				INSERT INTO #precedingElements
				SELECT L.startElementID
				FROM ASRSysWorkflowLinks L
				INNER JOIN ASRSysWorkflowElements E ON L.startElementID = E.ID
				WHERE L.endElementID = @piElementID
					AND E.type <> 9 -- 9 = Connector 2
			
				DECLARE precedingConnectorsCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.connectionPairID
				FROM ASRSysWorkflowLinks L
				INNER JOIN ASRSysWorkflowElements E ON L.startElementID = E.ID
				WHERE L.endElementID = @piElementID
					AND E.type = 9 -- 9 = Connector 2
			
				OPEN precedingConnectorsCursor
				FETCH NEXT FROM precedingConnectorsCursor INTO @iConnectorPairID
				WHILE (@@fetch_status = 0)
				BEGIN
					EXEC spASRGetPrecedingWorkflowElements @iConnectorPairID, @superCursor OUTPUT
			
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #precedingElements (elementID) VALUES (@iTemp)
						
						FETCH NEXT FROM @superCursor INTO @iTemp 
					END
					CLOSE @superCursor
					DEALLOCATE @superCursor
			
					FETCH NEXT FROM precedingConnectorsCursor INTO @iConnectorPairID
				END
				CLOSE precedingConnectorsCursor
				DEALLOCATE precedingConnectorsCursor
			
				/* Return the cursor of preceding elements. */
				SET @precedingElements = CURSOR FORWARD_ONLY STATIC FOR
					SELECT elementID 
					FROM #precedingElements
				OPEN @precedingElements
			
				DROP TABLE #precedingElements
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetStoredDataActionDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRGetStoredDataActionDetails]

		EXEC('CREATE PROCEDURE dbo.spASRGetStoredDataActionDetails
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psSQL		varchar(8000)	OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iPersonnelTableID	integer,
					@iInitiatorID		integer,
					@iDataTableID		integer,
					@iDataRecord		integer,
					@iRecordID		integer,
					@iDataAction		integer,
					@sTableName		varchar(8000),
					@sIDColumnName	varchar(8000),
					@iColumnID		integer, 
					@sColumnName		varchar(8000), 
					@iColumnDataType	integer, 
					@sColumnList		varchar(8000),
					@sValueList		varchar(8000),
					@sValue		varchar(8000)
			
				SET @psSQL = ''''
				SET @iRecordID = 0
			
				SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel''
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				SELECT @iDataAction = dataAction,
					@iDataTableID = dataTableID,
					@iDataRecord = dataRecord
				FROM ASRSysWorkflowElements
				WHERE ID = @piElementID
			
				SELECT @sTableName = tableName
				FROM ASRSysTables
				WHERE tableID = @iDataTableID
			
				IF @iDataRecord = 0 -- 0 = Initiator''s record
				BEGIN
					SET @iRecordID = @iInitiatorID
			
					IF @iDataTableID = @iPersonnelTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iPersonnelTableID)
					END
				END
			
				IF @iDataAction = 0 OR @iDataAction = 1
				BEGIN
					/* INSERT or UPDATE. */
					SET @sColumnList = ''''
					SET @sValueList = ''''
			
					DECLARE columnCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT EC.columnID,
						SC.columnName,
						SC.dataType,
						CASE
							WHEN EC.valueType = 0 THEN EC.value
							ELSE (SELECT IV.value
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElements WE ON IV.elementID = WE.ID
								INNER JOIN ASRSysWorkflowElements WE2 ON WE.workflowID = WE2.workflowID
								WHERE WE.identifier = EC.WFFormIdentifier
									AND WE2.id = @piElementID
									AND IV.instanceID = @piInstanceID
									AND IV.identifier = EC.WFValueIdentifier)
						END AS [value]
					FROM ASRSysWorkflowElementColumns EC
					INNER JOIN ASRSysColumns SC ON EC.columnID = SC.columnID
					WHERE EC.elementID = @piElementID
			
					OPEN columnCursor
					FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @iDataAction = 0 
						BEGIN
							/* INSERT. */
							SET @sColumnList = @sColumnList
								+ CASE
									WHEN LEN(@sColumnList) > 0 THEN '',''
									ELSE ''''
								END
								+ @sColumnName
			
							SET @sValueList = @sValueList
								+ CASE
									WHEN LEN(@sValueList) > 0 THEN '',''
									ELSE ''''
								END
								+ CASE
									WHEN @iColumnDataType = 12 
										OR @iColumnDataType = 11 
										OR @iColumnDataType = -1 THEN '''''''' + @sValue + '''''''' -- 11 = date, 12 = varchar, -1 = working pattern
									ELSE @sValue -- integer, logic, numeric
								END
						END
						ELSE
						BEGIN
							/* UPDATE. */
							SET @sColumnList = @sColumnList
								+ CASE
									WHEN LEN(@sColumnList) > 0 THEN '',''
									ELSE ''''
								END
								+ @sColumnName
								+ '' = ''
								+ CASE
									WHEN @iColumnDataType = 12 
										OR @iColumnDataType = 11 
										OR @iColumnDataType = -1 THEN '''''''' + @sValue + '''''''' -- 11 = date, 12 = varchar, -1 = working pattern
									ELSE @sValue -- integer, logic, numeric
								END
						END
			
						FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue
					END
			
					CLOSE columnCursor
					DEALLOCATE columnCursor
			
					IF @iDataAction = 0 
					BEGIN
						/* INSERT. */
						SET @sColumnList = @sColumnList
							+ CASE
								WHEN LEN(@sColumnList) > 0 THEN '',''
								ELSE ''''
							END
							+ @sIDColumnName
			
						SET @sValueList = @sValueList
							+ CASE
								WHEN LEN(@sValueList) > 0 THEN '',''
								ELSE ''''
							END
							+ convert(varchar(8000), @iRecordID)
					END
			
					IF LEN(@sColumnList) > 0
					BEGIN
						IF @iDataAction = 0 
						BEGIN
							/* INSERT. */
							SET @psSQL = ''INSERT INTO '' + @sTableName
								+ '' ('' + @sColumnList + '')''
								+ '' VALUES('' + @sValueList + '')''
						END
						ELSE
						BEGIN
							/* UPDATE. */
							SET @psSQL = ''UPDATE '' + @sTableName
								+ '' SET '' + @sColumnList
								+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(8000), @iRecordID)
						END
					END
				END
			
				IF @iDataAction = 2
				BEGIN
					/* DELETE. */
					SET @psSQL = ''DELETE FROM '' + @sTableName
						+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(8000), @iRecordID)
				END	
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetSucceedingWorkflowElements]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRGetSucceedingWorkflowElements]

		EXEC('CREATE PROCEDURE dbo.spASRGetSucceedingWorkflowElements
			(
				@piElementID		integer,
				@succeedingElements	cursor varying output
			)
			AS
			BEGIN
				CREATE TABLE #succeedingElements (elementID integer)
			
				/* Return the cursor of succeeding elements. */
				SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
					SELECT elementID 
					FROM #succeedingElements
				OPEN @succeedingElements
			
				DROP TABLE #succeedingElements
			END')

		EXEC('ALTER PROCEDURE dbo.spASRGetSucceedingWorkflowElements
			(
				@piElementID		integer,
				@succeedingElements	cursor varying output
			)
			AS
			BEGIN
				/* Return the IDs of the workflow elements that succeed the given element.
				This ignores connection elements.
				NB. This does work for elements with multiple outbound flows. */
				DECLARE
					@iConnectorPairID	integer,
					@superCursor		cursor,
					@iTemp		integer
				
				CREATE TABLE #succeedingElements (elementID integer)
			
				/* Get the non-connector elements. */
				INSERT INTO #succeedingElements
				SELECT L.endElementID
				FROM ASRSysWorkflowLinks L
				INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
				WHERE L.startElementID = @piElementID
					AND E.type <> 8 -- 8 = Connector 1
			
				DECLARE succeedingConnectorsCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.connectionPairID
				FROM ASRSysWorkflowLinks L
				INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
				WHERE L.startElementID = @piElementID
					AND E.type = 8 -- 8 = Connector 1
			
				OPEN succeedingConnectorsCursor
				FETCH NEXT FROM succeedingConnectorsCursor INTO @iConnectorPairID
				WHILE (@@fetch_status = 0)
				BEGIN
					EXEC spASRGetSucceedingWorkflowElements @iConnectorPairID, @superCursor OUTPUT
			
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
						
						FETCH NEXT FROM @superCursor INTO @iTemp 
					END
					CLOSE @superCursor
					DEALLOCATE @superCursor
			
					FETCH NEXT FROM succeedingConnectorsCursor INTO @iConnectorPairID
				END
				CLOSE succeedingConnectorsCursor
				DEALLOCATE succeedingConnectorsCursor
			
				/* Return the cursor of succeeding elements. */
				SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
					SELECT elementID 
					FROM #succeedingElements
				OPEN @succeedingElements
			
				DROP TABLE #succeedingElements
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetWorkflowFormItems]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRGetWorkflowFormItems]

		EXEC('CREATE PROCEDURE dbo.spASRGetWorkflowFormItems
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psErrorMessage	varchar(8000)	OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iID			integer,
					@iItemType		integer,
					@iDBColumnID		integer,
					@iDBRecord		integer,
					@sWFFormIdentifier	varchar(8000),
					@sWFValueIdentifier	varchar(8000),
					@sValue		varchar(8000),
					@sSQL			nvarchar(4000),
					@sSQLParam		nvarchar(4000),
					@sTableName		sysname,
					@sColumnName		sysname,
					@iInitiatorID		integer,
					@iRecordID		integer,
					@iStatus		integer,
					@iCount		integer
			
				/* Check the given instance still exists. */
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				IF @iCount = 0
				BEGIN
					SET @psErrorMessage = ''This workflow step is invalid. The workflow process may have been completed.''
					RETURN
				END
			
				/* Check if the step has already been completed! */
				SELECT @iStatus = ASRSysWorkflowInstanceSteps.status
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
				IF @iStatus = 3
				BEGIN
					SET @psErrorMessage = ''This workflow step has already been completed.''
					RETURN
				END
			
				SET @psErrorMessage = ''''
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				CREATE TABLE #itemValues (ID integer, value varchar(8000))	
			
				DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowElementItems.ID,
					ASRSysWorkflowElementItems.itemType,
					ASRSysWorkflowElementItems.dbColumnID,
					ASRSysWorkflowElementItems.dbRecord,
					ASRSysWorkflowElementItems.wfFormIdentifier,
					ASRSysWorkflowElementItems.wfValueIdentifier
				FROM ASRSysWorkflowElementItems
				WHERE ASRSysWorkflowElementItems.elementID = @piElementID
					AND (ASRSysWorkflowElementItems.itemType = 1 OR ASRSysWorkflowElementItems.itemType = 4)
			
				OPEN itemCursor
				FETCH NEXT FROM itemCursor INTO @iID, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @iItemType = 1
					BEGIN
						/* Database value. */
						SELECT @sTableName = ASRSysTables.tableName, 
							@sColumnName = ASRSysColumns.columnName
						FROM ASRSysColumns
						INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
						WHERE ASRSysColumns.columnID = @iDBColumnID
			
						IF @iDBRecord = 0 SET @iRecordID = @iInitiatorID
			
						SET @sSQL = ''SELECT @sValue = '' + @sTableName + ''.'' + @sColumnName +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
						SET @sSQLParam = N''@sValue varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT
					END
					ELSE
					BEGIN
						/* Workflow value. */
						SELECT @sValue = ASRSysWorkflowInstanceValues.value
						FROM ASRSysWorkflowInstanceValues
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
							AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
							AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
					END
			
					INSERT INTO #itemValues (ID, value)
					VALUES (@iID, @sValue)
			
					FETCH NEXT FROM itemCursor INTO @iID, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
				END
				CLOSE itemCursor
				DEALLOCATE itemCursor
			
				SELECT ASRSysWorkflowElementItems.*, #itemValues.value
				FROM ASRSysWorkflowElementItems
				LEFT OUTER JOIN #itemValues ON ASRSysWorkflowElementItems.ID = #itemValues.ID
				WHERE ASRSysWorkflowElementItems.elementID = @piElementID
				ORDER BY ASRSysWorkflowElementItems.ID
			
				DROP TABLE #itemValues
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRInstantiateWorkflow]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRInstantiateWorkflow]

		EXEC('CREATE PROCEDURE dbo.spASRInstantiateWorkflow
			(
				@piWorkflowID		integer,			
				@piInstanceID		integer		OUTPUT,
				@psFormElements	varchar(8000)	OUTPUT
			)
			AS
			BEGIN
				DECLARE
					@iInitiatorID		integer,
					@iStepID		integer,
					@iElementID		integer,
					@iRecordID		integer,
					@iRecordCount		integer,
					@sSQL			nvarchar(4000),
					@hResult		integer
			
				SET @iInitiatorID = 0
				SET @psFormElements = ''''
			
				SET @sSQL = ''spASRSysGetCurrentUserRecordID''
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @hResult = 0
			
					EXEC @hResult = @sSQL 
						@iRecordID OUTPUT,
						@iRecordCount OUTPUT
				END
			
				IF NOT @iRecordID IS null SET @iInitiatorID = @iRecordID
			
				/* Create the Workflow Instance record, and remember the ID. */
				INSERT INTO ASRSysWorkflowInstances (workflowID, initiatorID)
				VALUES (@piWorkflowID, @iInitiatorID)
				
				SELECT @piInstanceID = MAX(id)
				FROM ASRSysWorkflowInstances
			
				/* Create the Workflow Instance Steps records. 
				Set the first steps'' status to be 1 (pending Workflow Engine action). 
				Set all subsequent steps'' status to be 0 (on hold). */
			
				INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime)
				SELECT 
					@piInstanceID, 
					ASRSysWorkflowElements.ID, 
					CASE
						WHEN ASRSysWorkflowElements.ID IN (SELECT ASRSysWorkflowLinks.endElementID
							FROM ASRSysWorkflowLinks
							INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowLinks.startElementID = ASRSysWorkflowElements.ID
							WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
								AND ASRSysWorkflowElements.type = 0) THEN 1
						ELSE 0
					END, 
					CASE
						WHEN ASRSysWorkflowElements.ID IN (SELECT ASRSysWorkflowLinks.endElementID
							FROM ASRSysWorkflowLinks
							INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowLinks.startElementID = ASRSysWorkflowElements.ID
							WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
								AND ASRSysWorkflowElements.type = 0) THEN getdate()
						ELSE null
					END
				FROM ASRSysWorkflowElements 
				WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID
					AND ASRSysWorkflowElements.type <> 0
					AND ASRSysWorkflowElements.type <> 1
			
				/* Create the Workflow Instance Value records. */
				INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
				SELECT @piInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElementItems.identifier
				FROM ASRSysWorkflowElementItems 
				INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowElements.type = 2
					AND (ASRSysWorkflowElementItems.itemType = 3 OR ASRSysWorkflowElementItems.itemType = 0)
			
				/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
				DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowInstanceSteps.ID,
					ASRSysWorkflowInstanceSteps.elementID
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND ASRSysWorkflowElements.type = 2
			
				OPEN formsCursor
				FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
				WHILE (@@fetch_status = 0) 
				BEGIN
					SET @psFormElements = @psFormElements + convert(varchar(8000), @iElementID) + char(9)
			
					/* Change the step''s status to be 2 (pending user input). */
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 2
					WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID
			
					FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
				END
				CLOSE formsCursor
				DEALLOCATE formsCursor
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRWorkflowActionFailed]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRWorkflowActionFailed]

		EXEC('CREATE PROCEDURE dbo.spASRWorkflowActionFailed
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psMessage		varchar(8000)
			)
			AS
			BEGIN
				UPDATE ASRSysWorkflowInstanceSteps
				SET status = 4,	-- 4 = failed
					message = @psMessage
				WHERE instanceID = @piInstanceID
					AND elementID = @piElementID
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRCancelPendingPrecedingWorkflowElements]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRCancelPendingPrecedingWorkflowElements]

		EXEC('CREATE PROCEDURE dbo.spASRCancelPendingPrecedingWorkflowElements
			(
				@piInstanceID			integer,
				@piElementID			integer
			)
			AS
			BEGIN
				PRINT ''dummy''
			END')

		EXEC('ALTER PROCEDURE dbo.spASRCancelPendingPrecedingWorkflowElements
			(
				@piInstanceID			integer,
				@piElementID			integer
			)
			AS
			BEGIN
				/* Cancel (ie. set status to 0 for all workflow pending (ie. status 1 or 2) elements that precede the given element.
				This ignores connection elements.
				NB. This does work for elements with multiple inbound flows. */
				DECLARE
					@iConnectorPairID	integer,
					@iElementID		integer,
					@iStepID		integer,
					@superCursor		cursor,
					@iTemp		integer
				
				CREATE TABLE #precedingElements (elementID integer)
			
				EXEC spASRGetPrecedingWorkflowElements @piElementID, @superCursor output
				
				FETCH NEXT FROM @superCursor INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #precedingElements (elementID) VALUES (@iTemp)
					
					FETCH NEXT FROM @superCursor INTO @iTemp 
				END
				CLOSE @superCursor
				DEALLOCATE @superCursor
			
				/* Return the recordset of preceding elements. */
				DECLARE elementsCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.elementID,
					S.ID
				FROM #precedingElements E
				INNER JOIN ASRSysWorkflowInstanceSteps S ON E.elementID = S.elementID
				WHERE S.instanceID = @piInstanceID
					AND (S.status = 1 OR S.status = 2)
			
				OPEN elementsCursor
				FETCH NEXT FROM elementsCursor INTO @iElementID, @iStepID
				WHILE (@@fetch_status = 0)
				BEGIN
					UPDATE ASRSysWorkflowInstanceSteps
					SET status = 0
					WHERE ID = @iStepID
			
					EXEC spASRCancelPendingPrecedingWorkflowElements @piInstanceID, @iElementID
			
					FETCH NEXT FROM elementsCursor INTO @iElementID, @iStepID
				END
				CLOSE elementsCursor
				DEALLOCATE elementsCursor
			
				DROP TABLE #precedingElements
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetWorkflowEmailMessage]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRGetWorkflowEmailMessage]

		EXEC('CREATE PROCEDURE dbo.spASRGetWorkflowEmailMessage
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psMessage		varchar(8000)	OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iInitiatorID		integer,
					@sCaption		varchar(8000),
					@iItemType		integer,
					@iDBColumnID		integer,
					@iDBRecord		integer,
					@sWFFormIdentifier	varchar(8000),
					@sWFValueIdentifier	varchar(8000),
					@sValue		varchar(8000),
					@sTableName		sysname,
					@sColumnName		sysname,
					@iRecordID		integer,
					@sSQL			nvarchar(4000),
					@sSQLParam		nvarchar(4000),
					@iCount		integer,
					@iElementID		integer,
					@superCursor		cursor,
					@iTemp		integer
			
				SET @psMessage = ''''
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT EI.caption,
					EI.itemType,
					EI.dbColumnID,
					EI.dbRecord,
					EI.wfFormIdentifier,
					EI.wfValueIdentifier
				FROM ASRSysWorkflowElementItems EI
				WHERE EI.elementID = @piElementID
				ORDER BY EI.ID
			
				OPEN itemCursor
				FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @iItemType = 1
					BEGIN
						/* Database value. */
						SELECT @sTableName = ASRSysTables.tableName, 
							@sColumnName = ASRSysColumns.columnName
						FROM ASRSysColumns
						INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
						WHERE ASRSysColumns.columnID = @iDBColumnID
			
						IF @iDBRecord = 0 SET @iRecordID = @iInitiatorID
			
						SET @sSQL = ''SELECT @sValue = '' + @sTableName + ''.'' + @sColumnName +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
						SET @sSQLParam = N''@sValue varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT
			
						IF @sValue IS null SET @sValue = ''''
			
						SET @psMessage = @psMessage
							+ @sValue
					END
			
					IF @iItemType = 2
					BEGIN
						/* Label value. */
						SET @psMessage = @psMessage
							+ @sCaption
					END
			
					IF @iItemType = 4
					BEGIN
						/* Workflow value. */
						SELECT @sValue = ASRSysWorkflowInstanceValues.value
						FROM ASRSysWorkflowInstanceValues
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
							AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
							AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
			
						IF @sValue IS null SET @sValue = ''''
			
						SET @psMessage = @psMessage
							+ @sValue
					END
			
					FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
				END
				CLOSE itemCursor
				DEALLOCATE itemCursor
			
				/* Append the link to the webform that follows this element (ignore connectors) if there are any. */
				CREATE TABLE #succeedingElements (elementID integer)
			
				EXEC spASRGetSucceedingWorkflowElements @piElementID, @superCursor OUTPUT
			
				FETCH NEXT FROM @superCursor INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
					
					FETCH NEXT FROM @superCursor INTO @iTemp 
				END
				CLOSE @superCursor
				DEALLOCATE @superCursor
			
				SELECT @iCount = COUNT(*)
				FROM #succeedingElements SE
				INNER JOIN ASRSysWorkflowElements WE ON SE.elementID = WE.id
				WHERE WE.type = 2 -- 2 = Web Form element
			
				IF @iCount > 0 
				BEGIN
					SET @psMessage = @psMessage + CHAR(13) + CHAR(13)
						+ ''Click on the following link''
						+ CASE
							WHEN @iCount = 1 THEN ''''
							ELSE ''s''
						END
						+ '' to action:''
						+ CHAR(13)
			
					DECLARE elementCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT SE.elementID, WE.caption
					FROM #succeedingElements SE
					INNER JOIN ASRSysWorkflowElements WE ON SE.elementID = WE.ID
				
					OPEN elementCursor
					FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
					WHILE (@@fetch_status = 0)
					BEGIN
						SET @psMessage = @psMessage + CHAR(13) +
							@sCaption + '' - http://mrgrumpy/hrproworkflow/?'' + convert(varchar(8000), @piInstanceID) + ''&'' + convert(varchar(8000), @iElementID)
			
						FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
					END
					CLOSE elementCursor
			
					DEALLOCATE elementCursor
				END
			
				DROP TABLE #succeedingElements
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRSubmitWorkflowStep]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRSubmitWorkflowStep]

		EXEC('CREATE PROCEDURE dbo.spASRSubmitWorkflowStep
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psFormInput1		varchar(8000),
				@psFormInput2		varchar(8000)
			)
			AS
			BEGIN
				DECLARE
					@iIndex1		integer,
					@iIndex2		integer,
					@iID			integer,
					@sID			varchar(8000),
					@sValue		varchar(8000),
					@iElementType		integer,
					@iPreviousElementID	integer,
					@iValue		integer,
					@hResult		integer,
					@sTo			varchar(8000),
					@sMessage		varchar(8000),
					@iEmailID		integer,
					@iEmailRecord		integer,
					@iEmailRecordID	integer,
					@sSQL			nvarchar(4000),
					@iCount		integer,
					@superCursor		cursor,
					@iTemp		integer
			
				/* Put the submitted form values into the ASRSysWorkflowInstanceValues table. */
				WHILE charindex(CHAR(9), @psFormInput1 + @psFormInput2) > 0
				BEGIN
					SET @iIndex1 = charindex(CHAR(9), @psFormInput1 + @psFormInput2)
					SET @iIndex2 = charindex(CHAR(9), @psFormInput1 + @psFormInput2, @iIndex1+1)
			
					SET @sID = replace(LEFT(@psFormInput1 + @psFormInput2, @iIndex1-1), '''''''', '''''''''''')
					SET @sValue = replace(SUBSTRING(@psFormInput1 + @psFormInput2, @iIndex1+1, @iIndex2-@iIndex1-1), '''''''', '''''''''''')
			
					UPDATE ASRSysWorkflowInstanceValues
					SET ASRSysWorkflowInstanceValues.value = @sValue
					WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceValues.elementID = @piElementID
						AND ASRSysWorkflowInstanceValues.identifier = (SELECT ASRSysWorkflowElementItems.identifier
							FROM ASRSysWorkflowElementItems
							WHERE ASRSysWorkflowElementItems.ID = convert(integer, @sID))
			
					IF @iIndex2 > len(@psFormInput1)
					BEGIN
						SET @iIndex2 = @iIndex2 - len(@psFormInput1)
						SET @psFormInput1 = ''''
						SET @psFormInput2 = SUBSTRING(@psFormInput2, @iIndex2+1, LEN(@psFormInput2) - @iIndex2)
					END
					ELSE
					BEGIN
						SET @psFormInput1 = SUBSTRING(@psFormInput1, @iIndex2+1, LEN(@psFormInput1) - @iIndex2)
					END
				END
			
				/* Get the type of the given element */
				SELECT @iElementType = E.type,
					@iEmailID = E.emailID,
					@iEmailRecord = E.emailRecord
				FROM ASRSysWorkflowElements E
				WHERE E.ID = @piElementID
			
				SET @hResult = 0
			
				IF @iElementType = 3 -- Email element
				BEGIN
					/* Get the email recipient. */
					SET @sTo = ''''
					SET @iEmailRecordID = 0
			
					SET @sSQL = ''spASRSysEmailAddr''
					IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
					BEGIN
						/* Get the record ID required. */
						IF @iEmailRecord = 0
						BEGIN
							/* Initiator''s record. */
							SELECT @iEmailRecordID = ASRSysWorkflowInstances.initiatorID
							FROM ASRSysWorkflowInstances
							WHERE ASRSysWorkflowInstances.ID = @piInstanceID
						END
			
						/* Get the recipient''s address. */
						EXEC @hResult = @sSQL @sTo OUTPUT, @iEmailID, @iEmailRecordID
					END
			
					/* Build the email message. */
					EXEC spASRGetWorkflowEmailMessage @piInstanceID, @piElementID, @sMessage OUTPUT
			
					/* Send the email. */
					EXEC spASRSendMail @hResult OUTPUT,
							@sTo,
							'''',
							'''',
							''HR Pro Workflow'',
							@sMessage,
							''''
				END
			
				IF @hResult = 0
				BEGIN
					/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 3,
						ASRSysWorkflowInstanceSteps.completionDateTime = getdate()
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
				
					IF @iElementType = 4 -- Decision element
					BEGIN
						/* Get the ID of the elements that precede the Decision element. */
						CREATE TABLE #precedingElements (elementID integer)
			
						EXEC spASRGetPrecedingWorkflowElements @piElementID, @superCursor OUTPUT
				
						FETCH NEXT FROM @superCursor INTO @iTemp
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO #precedingElements (elementID) VALUES (@iTemp)
							
							FETCH NEXT FROM @superCursor INTO @iTemp 
						END
						CLOSE @superCursor
						DEALLOCATE @superCursor
			
						SELECT TOP 1 @iPreviousElementID = elementID
						FROM #precedingElements
			
						DROP TABLE #precedingElements
					
						SELECT @iValue = convert(integer, IV.value)
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
						WHERE IV.elementID = @iPreviousElementID
							AND IV.instanceid = @piInstanceID
							AND E.ID = @piElementID
					
						IF @iValue IS null SET @iValue = 0
			
						CREATE TABLE #succeedingElements2 (elementID integer)
			
						EXEC spASRGetDecisionSucceedingWorkflowElements @piElementID, @iValue, @superCursor OUTPUT
			
						FETCH NEXT FROM @superCursor INTO @iTemp
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO #succeedingElements2 (elementID) VALUES (@iTemp)
							
							FETCH NEXT FROM @superCursor INTO @iTemp 
						END
						CLOSE @superCursor
						DEALLOCATE @superCursor
			
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate()
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT #succeedingElements2.elementID 
								FROM #succeedingElements2)
			
						DROP TABLE #succeedingElements2
					END
					ELSE
					BEGIN
						CREATE TABLE #succeedingElements (elementID integer)
			
						EXEC spASRGetSucceedingWorkflowElements @piElementID, @superCursor OUTPUT
			
						FETCH NEXT FROM @superCursor INTO @iTemp
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
							
							FETCH NEXT FROM @superCursor INTO @iTemp 
						END
						CLOSE @superCursor
						DEALLOCATE @superCursor
			
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate()
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT #succeedingElements.elementID
								FROM #succeedingElements)
			
						DROP TABLE #succeedingElements
					END
				
					/* Set activated Web Forms to be ''pending'' (to be done by the user) */
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 2
					WHERE ASRSysWorkflowInstanceSteps.id IN (
						SELECT ASRSysWorkflowInstanceSteps.ID
						FROM ASRSysWorkflowInstanceSteps
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowInstanceSteps.status = 1
							AND ASRSysWorkflowElements.type = 2)
			
					/* Remove all record if the workflow has completed.
					ie. if it has no steps with status 1 (activated) or 2 (pending user action) */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowInstanceSteps
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.status IN (1, 2)
			
					IF @iCount = 0 
					BEGIN
						UPDATE ASRSysWorkflowInstances
						SET ASRSysWorkflowInstances.completionDateTime = getdate()
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
						/* NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
						is performed by a DELETE trigger on the ASRSysWorkflowInstances table. */
					END
				END
			END')

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRActionActiveWorkflowSteps]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
		drop procedure [dbo].[spASRActionActiveWorkflowSteps]

		EXEC('CREATE PROCEDURE dbo.spASRActionActiveWorkflowSteps
			AS
			BEGIN
				/* Return a recordset of the workflow steps that need to be actioned by the Workflow service.
				Action any that can be actioned immediately. */
				DECLARE
					@iAction		integer, -- 0 = do nothing, 1 = submit step, 2 = change status to ''2'', 3 = Summing Junction check, 4 = Or check
					@iElementType		integer,
					@iInstanceID		integer,
					@iElementID		integer,
					@iStepID		integer,
					@iCount		integer,
					@sStatus		bit,
					@sMessage		varchar(8000),
					@superCursor		cursor,
					@iTemp		integer
			
				DECLARE stepsCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.type,
					S.instanceID,
					E.ID,
					S.ID
				FROM ASRSysWorkflowInstanceSteps S
				INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
				WHERE S.status = 1
					AND E.type <> 5 -- 5 = StoredData elements handled in the service
			
				OPEN stepsCursor
				FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID
				WHILE (@@fetch_status = 0)
				BEGIN
					SET @iAction = 
						CASE
							WHEN @iElementType = 1 THEN 1	-- Terminator
							WHEN @iElementType = 2 THEN 2	-- Web form (action required from user)
							WHEN @iElementType = 3 THEN 1	-- Email
							WHEN @iElementType = 4 THEN 1	-- Decision
							WHEN @iElementType = 6 THEN 3	-- Summing Junction
							WHEN @iElementType = 7 THEN 4	-- Or	
							WHEN @iElementType = 8 THEN 1	-- Connector 1
							WHEN @iElementType = 9 THEN 1	-- Connector 2
							ELSE 0					-- Unknown
						END
					
					IF @iAction = 3 -- Summing Junction check
					BEGIN
						/* Check if all preceding steps have completed before submitting this step. */
						CREATE TABLE #precedingElements (elementID integer)
					
						EXEC spASRGetPrecedingWorkflowElements @iElementID, @superCursor OUTPUT
				
						FETCH NEXT FROM @superCursor INTO @iTemp
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO #precedingElements (elementID) VALUES (@iTemp)
							
							FETCH NEXT FROM @superCursor INTO @iTemp 
						END
						CLOSE @superCursor
						DEALLOCATE @superCursor
			
						SELECT @iCount = COUNT(*)
						FROM ASRSysWorkflowInstanceSteps WIS
						INNER JOIN #precedingElements PE ON WIS.elementID = PE.elementID
						WHERE WIS.instanceID = @iInstanceID
							AND WIS.status <> 3 -- 3 = completed
			
						/* If all preceding steps have been completed submit the Summing Junction step. */
						IF @iCount = 0 SET @iAction = 1
			
						DROP TABLE #precedingElements
					END
			
					IF @iAction = 4 -- Or check
					BEGIN
						/* Check if any preceding steps have completed before submitting this step. */
						CREATE TABLE #precedingElements2 (elementID integer)
			
						EXEC spASRGetPrecedingWorkflowElements @iElementID, @superCursor output
			
						FETCH NEXT FROM @superCursor INTO @iTemp
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO #precedingElements2 (elementID) VALUES (@iTemp)
						
							FETCH NEXT FROM @superCursor INTO @iTemp 
						END
						CLOSE @superCursor
						DEALLOCATE @superCursor
			
						SELECT @iCount = COUNT(*)
						FROM ASRSysWorkflowInstanceSteps WIS
						INNER JOIN #precedingElements2 PE ON WIS.elementID = PE.elementID
						WHERE WIS.instanceID = @iInstanceID
							AND WIS.status = 3 -- 3 = completed
			
						/* If all preceding steps have been completed submit the Or step. */
						IF @iCount > 0 
						BEGIN
							/* Cancel any preceding steps that are not completed as they are no longer required. */
							EXEC spASRCancelPendingPrecedingWorkflowElements @iInstanceID, @iElementID
			
							SET @iAction = 1
						END
			
						DROP TABLE #precedingElements2
					END
			
					IF @iAction = 1
					BEGIN
						EXEC spASRSubmitWorkflowStep @iInstanceID, @iElementID, '''', ''''
					END
			
					IF @iAction = 2
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET status = 2
						WHERE id = @iStepID
					END
			
					FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID
				END
			END')



/* ------------------------------------------------------------- */
PRINT 'Step X of X - Adding trigger to Workflow Instances'

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DEL_ASRSysWorkflowInstances]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[DEL_ASRSysWorkflowInstances]

		EXEC('CREATE TRIGGER DEL_ASRSysWorkflowInstances ON [dbo].[ASRSysWorkflowInstances] 
			FOR DELETE
			AS
			BEGIN
				/* Delete related records. */
				DELETE FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID IN (SELECT id FROM deleted)
			
				DELETE FROM ASRSysWorkflowInstanceValues
				WHERE ASRSysWorkflowInstanceValues.instanceID IN (SELECT id FROM deleted)
			END')

/* ------------------------------------------------------------- */
PRINT 'Step X of X - Update Support Email Address'

	UPDATE ASRSysSystemSettings
	Set [SettingValue] = 'support@asr.co.uk'
	WHERE [SettingValue] = 'helpdesk@asr.co.uk'


/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step X of X - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '2.20')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '2.20.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v2.20')


SELECT @NVarCommand = 'USE master
GRANT ALL ON master..xp_LoginConfig TO public
GRANT ALL ON master..xp_EnumGroups TO public
GRANT ALL ON master..xp_StartMail TO public
GRANT ALL ON master..xp_SendMail TO public'
EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'USE '+@DBName
EXEC sp_executesql @NVarCommand


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'refreshstoredprocedures'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'refreshstoredprocedures', 1)


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v2.20 Of HR Pro'
