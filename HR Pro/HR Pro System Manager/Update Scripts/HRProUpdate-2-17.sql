
/* ----------------------------------------------------- */
/* Update the database from version 2.16 to version 2.17 */
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
	@sSQLVersion nvarchar(20)

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

/* Exit if the database is not version 2.16 or 2.17. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '2.16') and (@sDBVersion <> '2.17')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
PRINT 'Step 1 of 22 - Creating new objects for Outlook Calendar Interface'

	if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysOutlookFolders]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysOutlookFolders] (
								[FolderID] [int] NULL ,
								[TableID] [int] NULL ,
								[Name] [varchar] (50) NULL ,
								[FolderType] [int] NULL ,
								[FixedPath] [varchar] (2000) NULL ,
								[ExprID] [int] NULL 
							) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysOutlookLinks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysOutlookLinks] (
								[LinkID] [int] NULL ,
								[TableID] [int] NULL ,
								[Title] [varchar] (50) NULL ,
								[FilterID] [int] NULL ,
								[BusyStatus] [int] NULL ,
								[StartDate] [int] NULL ,
								[EndDate] [int] NULL ,
								[TimeRange] [int] NULL ,
								[FixedStartTime] [varchar] (5) NULL ,
								[FixedEndTime] [varchar] (5) NULL ,
								[ColumnStartTime] [int] NULL ,
								[ColumnEndTime] [int] NULL ,
								[Subject] [int] NULL ,
								[Content] [text] NULL ,
								[Reminder] [bit] NULL ,
								[ReminderOffset] [int] NULL ,
								[ReminderPeriod] [int] NULL 
							) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysOutlookLinksColumns]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysOutlookLinksColumns] (
								[LinkID] [int] NULL ,
								[ColumnID] [int] NULL ,
								[Heading] [varchar] (50) NULL ,
								[Sequence] [int] NULL 
							) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysOutlookLinksDestinations]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysOutlookLinksDestinations] (
								[LinkID] [int] NULL ,
								[FolderID] [int] NULL 
							) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysOutlookEvents]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysOutlookEvents] (
								[LinkID] [int] NULL ,
								[FolderID] [int] NULL ,
								[TableID] [int] NULL ,
								[RecordID] [int] NULL ,
								[Refresh] [bit] NULL ,
								[Deleted] [bit] NULL ,
								[ErrorMessage] [varchar] (2000) NULL ,
								[StoreID] [varchar] (2000) NULL ,
								[EntryID] [varchar] (2000) NULL 
							) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	if not exists (select * from sysobjects where id = object_id(N'[dbo].[ASRSysOutlookEvents]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysOutlookEvents] (
								[LinkID] [int] NULL ,
								[FolderID] [int] NULL ,
								[TableID] [int] NULL ,
								[RecordID] [int] NULL ,
								[Refresh] [bit] NULL ,
								[Deleted] [bit] NULL ,
								[ErrorMessage] [varchar] (2000) NULL ,
								[StoreID] [varchar] (2000) NULL ,
								[EntryID] [varchar] (2000) NULL 
							) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysOutlookEvents')
	and name = 'Folder'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysOutlookEvents ADD 
					[Folder] [varchar] (255) NULL,
					[Subject] [varchar] (255) NULL,
					[StartDate] [datetime] NULL,
					[EndDate] [datetime] NULL'
		EXEC sp_executesql @NVarCommand
	END


	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysOutlookEvents')
	and name = 'RefreshDate'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysOutlookEvents ADD 
					[RefreshDate] [datetime] NULL'
		EXEC sp_executesql @NVarCommand
	END


	if not exists (select * from sysobjects where id = object_id(N'[dbo].[DEL_ASRSysOutlookLinks]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysOutlookLinks ON [dbo].[ASRSysOutlookLinks] 
							FOR DELETE 
							AS
							BEGIN
								DELETE FROM ASRSysOutlookLinksColumns WHERE LinkID NOT IN (SELECT LinkID FROM ASRSysOutlookLinks)
								DELETE FROM ASRSysOutlookLinksDestinations WHERE LinkID NOT IN (SELECT LinkID FROM ASRSysOutlookLinks)
							END'
		EXEC sp_executesql @NVarCommand
	END



	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASROutlookEventRefresh]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASROutlookEventRefresh]

	SELECT @NVarCommand = 'CREATE PROCEDURE dbo.spASROutlookEventRefresh
		(@LinkID int, @FolderID int, @TableID int, @RecordID int)
		 AS
		BEGIN
			IF EXISTS(SELECT * FROM ASRSysOutlookEvents WHERE LinkID = @LinkID AND FolderID = @FolderID AND TableID = @TableID AND RecordID = @RecordID)
			  UPDATE ASRSysOutlookEvents SET Refresh = 1 WHERE LinkID = @LinkID AND FolderID = @FolderID AND TableID = @TableID AND RecordID = @RecordID
			ELSE
			  INSERT ASRSysOutlookEvents(LinkID, FolderID, TableID, RecordID, Refresh, Deleted) VALUES (@LinkID,@FolderID, @TableID, @RecordID, 1, 0)
		END'

	EXEC sp_executesql @NVarCommand


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
						SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+convert(varchar(8000),isnull([''+@ColumnName+''],''''''''),''+@DateFormat+'') FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(4000),@RecordID)
					IF @DataType <> 11 AND @DataType <> 12
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
PRINT 'Step 2 of 22 - Adding Email Notifications to Batch Jobs'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysBatchJobName')
	and name = 'EmailFailed'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysBatchJobName ADD
					[EmailFailed] [int] NULL,
					[EmailSuccess] [int] NULL'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'UPDATE ASRSysBatchJobName SET
					[EmailFailed] = 0, [EmailSuccess] = 0'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 3 of 22 - Windows Authentication validate login user'

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRCheckNTLogin]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spASRCheckNTLogin]


	SELECT @NVarCommand = 'CREATE PROCEDURE spASRCheckNTLogin 
		(
		    @sLoginName varchar(800)
		)
		AS 
		BEGIN
		
			DECLARE @hResult integer
			DECLARE @strFoundName varchar(800)
			DECLARE @bFound bit
			
			SELECT @strFoundName = name from master..syslogins where name = @sLoginName and isntname = 1
		
			SET @bFound = 0	
			IF (@strFoundName IS NULL)
				BEGIN
					EXEC @hResult = sp_grantlogin @sLoginName
		
					IF @hResult = 0
						BEGIN
							EXEC sp_revokelogin @sLoginName
							SET @bFound = 1
						END
					ELSE
						BEGIN
							SET @bFound = 0
						END
		
				END
			ELSE 
				SET @bFound = 1
		
			SELECT @bFound
		
		END'

	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 4 of 22 - Delete obsolete "Field Entry Validation" Expressions.'

	DELETE FROM ASRSysExpressions WHERE TYPE = 2


/* ------------------------------------------------------------- */
PRINT 'Step 5 of 22 - Batch Jobs - Amending text on pause parameters'

	UPDATE [ASRSysBatchJobDetails] 
	SET [ASRSysBatchJobDetails].[Parameter] = 'N/A' 
	WHERE LOWER([ASRSysBatchJobDetails].[JobType]) <> '-- pause --' 


/* ------------------------------------------------------------- */
PRINT 'Step 6 of 22 - Adding embedded column type.'

	/* Add Embedded option column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysColumns')
	and name = 'Embedded'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysColumns ADD 
					[Embedded] [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysColumns SET [Embedded] = 0'
		EXEC sp_executesql @NVarCommand

	END

/* ------------------------------------------------------------- */
PRINT 'Step 7 of 22 - Adding maximum ole size column type.'

	/* Add Embedded option column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysColumns')
	and name = 'OLEType'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysColumns ADD 
					[OLEType] [int] NULL,
					[MaxOLESizeEnabled] [bit] NULL,
					[MaxOLESize] [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysColumns SET [MaxOLESizeEnabled] = 0, [MaxOLESize] = 0'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysColumns SET [OLEType] = [OLEOnServer]'
		EXEC sp_executesql @NVarCommand

	END

/* ------------------------------------------------------------- */
PRINT 'Step 8 of 22 - Amending Screen Controls Procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRGetControlDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRGetControlDetails]


	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].sp_ASRGetControlDetails 
		(
			@piScreenID int
		)
		AS
		BEGIN
			SELECT ASRSysControls.*, 
				ASRSysColumns.columnName, 
				ASRSysColumns.columnType, 
				ASRSysColumns.datatype,
				ASRSysColumns.defaultValue,
				ASRSysColumns.size, 
				ASRSysColumns.decimals, 
				ASRSysColumns.lookupTableID, 
				ASRSysColumns.lookupColumnID, 
				ASRSysColumns.lookupFilterColumnID, 
				ASRSysColumns.lookupFilterOperator, 
				ASRSysColumns.lookupFilterValueID, 
				ASRSysColumns.spinnerMinimum, 
				ASRSysColumns.spinnerMaximum, 
				ASRSysColumns.spinnerIncrement, 
				ASRSysColumns.mandatory, 
				ASRSysColumns.uniquecheck,
				ASRSysColumns.convertcase, 
				ASRSysColumns.mask, 
				ASRSysColumns.blankIfZero, 
				ASRSysColumns.multiline, 
				ASRSysColumns.alignment AS colAlignment, 
				ASRSysColumns.calcExprID, 
				ASRSysColumns.gotFocusExprID, 
				ASRSysColumns.lostFocusExprID, 
				ASRSysColumns.dfltValueExprID, 
				ASRSysColumns.calcTrigger, 
				ASRSysColumns.readOnly, 
				ASRSysColumns.statusBarMessage, 
				ASRSysColumns.errorMessage, 
				ASRSysColumns.linkTableID,
				ASRSysColumns.linkViewID,
				ASRSysColumns.linkOrderID,
				ASRSysColumns.Afdenabled,
				ASRSysTables.TableName,
				ASRSysColumns.Trimming,
				ASRSysColumns.Use1000Separator,
				ASRSysColumns.QAddressEnabled,
				ASRSysColumns.OLEType,
				ASRSysColumns.MaxOLESizeEnabled,
				ASRSysColumns.MaxOLESize
			FROM ASRSysControls 
			LEFT OUTER JOIN ASRSysTables 
				ON ASRSysControls.tableID = ASRSysTables.tableID 
			LEFT OUTER JOIN ASRSysColumns 
				ON ASRSysColumns.tableID = ASRSysControls.tableID 
					AND ASRSysColumns.columnID = ASRSysControls.columnID
			WHERE ASRSysControls.ScreenID = @piScreenID
			ORDER BY ASRSysControls.PageNo, 
				ASRSysControls.ControlLevel DESC, 
				ASRSysControls.tabIndex
		END'
	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 9 of 22 - Removing obsolete user email procedures'

	DECLARE HRProCursor CURSOR
	FOR SELECT sysusers.name+'.'+sysobjects.name
	FROM sysobjects
	JOIN sysusers ON sysusers.uid = sysobjects.uid
	WHERE sysobjects.name = 'spASRSysEmailAddr' and sysusers.name <> 'dbo'

	OPEN HRProCursor
	FETCH NEXT FROM HRProCursor INTO @sName
	WHILE @@FETCH_STATUS = 0
	BEGIN
		SELECT @NVarCommand = 'DROP PROCEDURE '+@sName
		EXECUTE sp_sqlexec @NVarCommand
		FETCH NEXT FROM HRProCursor INTO @sName
	END

	CLOSE HRProCursor
	DEALLOCATE HRProCursor

/* ------------------------------------------------------------- */
PRINT 'Step 10 of 22 - Updating Self-service Intranet Link table'

	SELECT @iRecCount = count(id) FROM syscolumns
	WHERE id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
		AND name = 'startMode'
	
	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
					[startMode] [int] NULL'
		EXEC sp_executesql @NVarCommand
	
		/* 
		new = 1
		first = 2
		find = 3
		*/
		SET @NVarCommand = 'UPDATE ASRSysSSIntranetLinks
			SET [startMode] = 
				CASE
					WHEN ASRSysSSIntranetLinks.screenID > 0 THEN
						CASE
							WHEN ASRSysTables.tableType = 1 THEN 2
							ELSE 1
						END
					ELSE 0
				END
			FROM ASRSysSSIntranetLinks
			LEFT OUTER JOIN ASRSysScreens ON ASRSysSSIntranetLinks.screenID = ASRSysScreens.screenID
			LEFT OUTER JOIN ASRSysTables ON ASRSysScreens.tableID = ASRSysTables.tableID'
	
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	WHERE id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
		AND name = 'selfServiceAccess'
	
	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
					[selfServiceAccess] [bit] NOT NULL DEFAULT 1'
		EXEC sp_executesql @NVarCommand
	
		SET @NVarCommand = 'UPDATE ASRSysSSIntranetLinks
			SET [selfServiceAccess] = 1'
	
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	WHERE id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
		AND name = 'ID'
	
	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
					[id] [int] IDENTITY (1, 1) NOT NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	WHERE id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
		AND name = 'utilityType'
	
	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
					[utilityType] [int] NULL'
		EXEC sp_executesql @NVarCommand
	
		SET @NVarCommand = 'UPDATE ASRSysSSIntranetLinks
			SET [utilityType] = -1'
	
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	WHERE id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
		AND name = 'utilityID'
	
	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
					[utilityID] [int] NULL'
		EXEC sp_executesql @NVarCommand
	
		SET @NVarCommand = 'UPDATE ASRSysSSIntranetLinks
			SET [utilityID] = 0'
	
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 11 of 22 - Creating Self-service Intranet stored procedures'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRIntGetDefaultOrder]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRIntGetDefaultOrder]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRIntGetLinkInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRIntGetLinkInfo]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRIntGetLinks]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRIntGetLinks]

	SELECT @NVarCommand = 'CREATE PROCEDURE spASRIntGetDefaultOrder (
		@piTableID	integer,
		@piOrderID	integer	OUTPUT
	)
	AS
	BEGIN
		SELECT @piOrderID = defaultOrderID
		FROM ASRSysTables
		WHERE tableID = @piTableID
	END'

	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'CREATE PROCEDURE spASRIntGetLinkInfo 
	(
		@piLinkID 		integer,
		@piScreenID		integer	OUTPUT,
		@piTableID		integer	OUTPUT,
		@psTitle		varchar(8000)	OUTPUT,
		@piStartMode		integer	OUTPUT,
		@pfSelfService		bit	OUTPUT
	)
	AS
	BEGIN
		SELECT 
			@piScreenID = ASRSysSSIntranetLinks.screenID,
			@piTableID = ASRSysScreens.tableID,
			@psTitle = ASRSysSSIntranetLinks.pageTitle,
			@piStartMode = ASRSysSSIntranetLinks.startMode,
			@pfSelfService = ASRSysSSIntranetLinks.selfServiceAccess
		FROM ASRSysSSIntranetLinks
		INNER JOIN ASRSysScreens ON ASRSysSSIntranetLinks.screenID = ASRSysScreens.screenID
		WHERE ID = @piLinkID
	END'

	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'CREATE PROCEDURE spASRIntGetLinks 
	(
		@piLinkType 		integer,
		@pfSelfServiceAccess	bit,
		@pfFullLinksExist	bit	OUTPUT
	)
	AS
	BEGIN
		DECLARE @iCount	integer
	
		SELECT ASRSysSSIntranetLinks.*
		FROM ASRSysSSIntranetLinks
		WHERE linkType = @piLinkType
			AND selfServiceAccess = @pfSelfServiceAccess
		ORDER BY linkOrder
	
		SET @pfFullLinksExist = 0
	
		IF @piLinkType = 1 AND @pfSelfServiceAccess = 1
		BEGIN
			SELECT @iCount = COUNT(*)
			FROM ASRSysSSIntranetLinks
			WHERE (linkType = 1
				OR linkType = 2)
				AND selfServiceAccess = 0
	
			IF @iCount > 0
			BEGIN
				SET @pfFullLinksExist = 1
			END
		END
	END'

	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 12 of 22 - Reset reduntant settings on OLE columns.'

	SET @NVarCommand = 'UPDATE ASRSysColumns SET [ReadOnly] = 0, [Mandatory] = 0, [Duplicate] = 0, [Audit] = 0, [ChildUniqueCheck] = 0, [UniqueCheck] = 0
		WHERE [dataType] = -4 OR [DataType] = -3'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 13 of 22 - Adding Auto Update Lookup Value column type.'

	/* Add Embedded option column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysColumns')
	and name = 'AutoUpdateLookupValues'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysColumns ADD 
					[AutoUpdateLookupValues] [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysColumns SET [AutoUpdateLookupValues] = 0'
		EXEC sp_executesql @NVarCommand

	END


/* ------------------------------------------------------------- */
PRINT 'Step 14 of 22 - Adding Tidy Up Windows Orphans stored procedure.'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRDeleteInvalidLogins]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRDeleteInvalidLogins]

	SELECT @NVarCommand = '-- This function must be marked as a system object otherwise the get_sid function doesn''t work. Microsoft don''t tell you this!!!
	--EXEC sp_MS_marksystemobject ''spASRDeleteInvalidLogins''
	CREATE PROCEDURE spASRDeleteInvalidLogins
	(
	    @pstrDomainName varchar(100)
	)
	AS
	
		DECLARE @cursLogins cursor
		DECLARE @loginName nvarchar(200)

		-- Are we privileged enough to run this script
		IF IS_SRVROLEMEMBER(''securityadmin'') = 0 and IS_SRVROLEMEMBER(''sysadmin'') = 0
		BEGIN
			RETURN 0
		END

		-- Lets get the invalid accounts into a swish little cursor
		SET @cursLogins = CURSOR LOCAL FAST_FORWARD FOR
		SELECT loginName from master.dbo.syslogins
			WHERE isntname = 1
			AND loginname like @pstrDomainName + ''\%''
			AND ((sid <> SUSER_SID(loginname) and SUSER_SID(loginname) is not null)	OR SUSER_SID(loginname) is null)

		-- Now lets get rid of the invalid accounts
		OPEN @cursLogins
		FETCH NEXT FROM @cursLogins INTO @LoginName
		WHILE (@@fetch_status = 0)
		BEGIN
			EXEC sp_revokelogin @LoginName
			FETCH NEXT FROM @cursLogins INTO @LoginName
		END
	
		-- Tidy Up
		CLOSE @cursLogins 	
		DEALLOCATE @cursLogins

		RETURN 0'

	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 15 of 22 - Amending Screen Control Details Procedure.'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRGetControlDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRGetControlDetails]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].sp_ASRGetControlDetails 
		(
			@piScreenID int
		)
		AS
		BEGIN
			SELECT ASRSysControls.*, 
				ASRSysColumns.columnName, 
				ASRSysColumns.columnType, 
				ASRSysColumns.datatype,
				ASRSysColumns.defaultValue,
				ASRSysColumns.size, 
				ASRSysColumns.decimals, 
				ASRSysColumns.lookupTableID, 
				ASRSysColumns.lookupColumnID, 
				ASRSysColumns.lookupFilterColumnID, 
				ASRSysColumns.lookupFilterOperator, 
				ASRSysColumns.lookupFilterValueID, 
				ASRSysColumns.spinnerMinimum, 
				ASRSysColumns.spinnerMaximum, 
				ASRSysColumns.spinnerIncrement, 
				ASRSysColumns.mandatory, 
				ASRSysColumns.uniquecheck,
				ASRSysColumns.uniquechecktype,
				ASRSysColumns.convertcase, 
				ASRSysColumns.mask, 
				ASRSysColumns.blankIfZero, 
				ASRSysColumns.multiline, 
				ASRSysColumns.alignment AS colAlignment, 
				ASRSysColumns.calcExprID, 
				ASRSysColumns.gotFocusExprID, 
				ASRSysColumns.lostFocusExprID, 
				ASRSysColumns.dfltValueExprID, 
				ASRSysColumns.calcTrigger, 
				ASRSysColumns.readOnly, 
				ASRSysColumns.statusBarMessage, 
				ASRSysColumns.errorMessage, 
				ASRSysColumns.linkTableID,
				ASRSysColumns.linkViewID,
				ASRSysColumns.linkOrderID,
				ASRSysColumns.Afdenabled,
				ASRSysTables.TableName,
				ASRSysColumns.Trimming,
				ASRSysColumns.Use1000Separator,
				ASRSysColumns.QAddressEnabled,
				ASRSysColumns.OLEType,
				ASRSysColumns.MaxOLESizeEnabled,
				ASRSysColumns.MaxOLESize,
				ASRSysColumns.AutoUpdateLookupValues			
			FROM ASRSysControls 
			LEFT OUTER JOIN ASRSysTables 
				ON ASRSysControls.tableID = ASRSysTables.tableID 
			LEFT OUTER JOIN ASRSysColumns 
				ON ASRSysColumns.tableID = ASRSysControls.tableID 
					AND ASRSysColumns.columnID = ASRSysControls.columnID
			WHERE ASRSysControls.ScreenID = @piScreenID
			ORDER BY ASRSysControls.PageNo, 
				ASRSysControls.ControlLevel DESC, 
				ASRSysControls.tabIndex
		END'

	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 16 of 22 - Removing email links from Link/OLE/Photo Columns.'

	DELETE FROM ASRSysEmailLinks
	WHERE ColumnID IN 
	(SELECT ColumnID FROM ASRSysColumns WHERE ASRSysColumns.ColumnType = 4 OR ASRSysColumns.DataType = -3 OR ASRSysColumns.DataType = -4)



/* ------------------------------------------------------------- */
PRINT 'Step 17 of 22 - Amending Bradford Factor Output Options.'

	--Remove Bradford output settings if CSV format.
	if exists(select * from ASRSysSystemSettings where section = 'bradfordfactor' and settingkey = 'format' and settingvalue = 1)
	begin
	  delete from ASRSysSystemSettings where [Section] = 'bradfordfactor' and
	  settingkey in ('format', 'preview', 'screen', 'save', 'saveexisting', 'filename', 'email', 'emailaddr', 'emailattachas', 'emailsubject')

	  insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
	  values('bradfordfactor', 'screen', 1)
	end


/* ------------------------------------------------------------- */
PRINT 'Step 18 of 22 - Updating Working Pattern Formatting.'

	--Right Trim all working patterns
	UPDATE ASRSysColumns SET Trimming = 3 WHERE DataType = -1

/* ------------------------------------------------------------- */
PRINT 'Step 19 of 22 - Table OLE stats stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRSysTableOLEStats]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRSysTableOLEStats]

	SELECT @NVarCommand = 'CREATE PROCEDURE dbo.spASRSysTableOLEStats
		(
		@pstrTableID int
		)
	AS
	BEGIN
	
		declare @strColumn nvarchar(200)
		declare @cmdstr varchar(4000)
		declare @strTableName varchar(100)
		
		create table #TempTable 
		(	[ColumnName] varchar(200),
			[EmbeddedSize] float)
	
		set @strTableName = (select tablename from asrSysTables where tableid = @pstrTableID)
		
		declare tabColumns cursor for select ColumnName from asrSysColumns where tableid = @pstrTableID and oletype = 2
		open tabColumns
		
			fetch next from tabColumns into @strColumn
		
			while @@fetch_status = 0
			begin
				
				set @cmdstr = ''select '''''' + @strColumn + '''''' as Name, sum(datalength('' + @strColumn + ''))'' + '' from '' + @strTableName
					+ '' where datalength('' + @strColumn + '') > 300''
				insert into #TempTable exec(@cmdstr)
				fetch next from tabColumns into @strColumn
			end
		
		select * from #TempTable
		
		close tabColumns
		deallocate tabColumns
		
		drop table #TempTable
		
	
	END'
	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
PRINT 'Step 20 of 22 - Cleanup database stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRDropTempObjects]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRDropTempObjects]

	SELECT @NVarCommand = 'CREATE PROCEDURE spASRDropTempObjects
		AS
		BEGIN

			DECLARE	@sObjectName varchar(2000),
					@sUsername varchar(2000),
					@sXType varchar(50)
						
			DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
			SELECT [dbo].[sysobjects].[name], [dbo].[sysusers].[name], [dbo].[sysobjects].[xtype]
			FROM [dbo].[sysobjects] 
					INNER JOIN [dbo].[sysusers]
					ON [dbo].[sysobjects].[uid] = [dbo].[sysusers].[uid]
			WHERE LOWER([dbo].[sysusers].[name]) != ''dbo''
					AND (OBJECTPROPERTY(id, N''IsUserTable'') = 1
						OR OBJECTPROPERTY(id, N''IsProcedure'') = 1
						OR OBJECTPROPERTY(id, N''IsInlineFunction'') = 1
						OR OBJECTPROPERTY(id, N''IsScalarFunction'') = 1
						OR OBJECTPROPERTY(id, N''IsTableFunction'') = 1)

			OPEN tempObjects
			FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType
			WHILE (@@fetch_status <> -1)
			BEGIN		
				IF UPPER(@sXType) = ''U''
					-- user table
					BEGIN
						EXEC (''DROP TABLE ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END

				IF UPPER(@sXType) = ''P''
					-- procedure
					BEGIN
						EXEC (''DROP PROCEDURE ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END

				IF UPPER(@sXType) = ''TF''
					-- UDF
					BEGIN
						EXEC (''DROP FUNCTION ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END

				IF UPPER(@sXType) = ''FN''
					-- UDF
					BEGIN
						EXEC (''DROP FUNCTION ['' + @sUsername + ''].['' + @sObjectName + '']'')
					END
				
				FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType
				
			END
			CLOSE tempObjects
			DEALLOCATE tempObjects
			
			EXEC (''DELETE FROM [dbo].[ASRSysSQLObjects]'')

		END'
	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 21 of 22 - Adding default to Columns table (KB000155)'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DF_ASRSysColumns_QAddressEnabled]') and OBJECTPROPERTY(id, N'IsConstraint') = 1)
	BEGIN
		ALTER TABLE ASRSysColumns
			DROP CONSTRAINT DF_ASRSysColumns_QAddressEnabled
	END

	ALTER TABLE ASRSysColumns ADD CONSTRAINT
		DF_ASRSysColumns_QAddressEnabled DEFAULT (0) FOR QAddressEnabled
	
	UPDATE ASRSysColumns SET QAddressEnabled = 0 WHERE QAddressEnabled IS NULL



/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step 22 of 22 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '2.17')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '2.17.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v2.17')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v2.17 Of HR Pro'
