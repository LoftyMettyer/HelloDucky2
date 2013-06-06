
/* ----------------------------------------------------- */
/* Update the database from version 2.18 to version 2.19 */
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
IF (@sDBVersion <> '2.18') and (@sDBVersion <> '2.19')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
PRINT 'Step 1 of 13 - Adding new columns to Email Queue'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysEmailQueue')
		and name = 'RepTo'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE ASRSysEmailQueue ADD
			                       [RepTo] [varchar] (4000) NULL,
			                       [RepCC] [varchar] (4000) NULL,
			                       [RepBCC] [varchar] (4000) NULL,
			                       [MsgText] [varchar] (8000) NULL,
			                       [Subject] [varchar] (4000) NULL,
			                       [Attachment] [varchar] (4000) NULL'
			EXEC sp_executesql @NVarCommand
		END


/* ------------------------------------------------------------- */
PRINT 'Step 2 of 13 - Creating new email stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRSendMail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRSendMail]

	SELECT @NVarCommand = 'CREATE PROCEDURE dbo.spASRSendMail(
				@hResult int OUTPUT,
				@To varchar(8000),
				@CC varchar(8000),
				@BCC varchar(8000),
				@Subject varchar(8000),
				@Message varchar(8000),
				@Attachment varchar(8000))
			AS
			BEGIN
				EXEC @hResult = master..xp_sendmail
					@recipients=@To,
					@copy_recipients=@CC,
					@blind_copy_recipients=@BCC,
					@subject=@Subject,
					@message=@Message,
					@attachments=@Attachment
			END'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 3 of 13 - Amending existing email stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASREmailImmediate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASREmailImmediate]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].spASREmailImmediate(@Username varchar(255)) AS
		BEGIN
			DECLARE @QueueID int,
				@LinkID int,
				@RecordID int,
				@ColumnID int,
				@ColumnValue varchar(8000),
				@RecDescID int,
				@RecDesc varchar(4000),
				@sSQL nvarchar(4000),
				@EmailDate datetime,
				@DateDue datetime,
				@hResult int,
				@blnEnabled int,
				@RecalculateRecordDesc bit,
				@TableID int
			DECLARE @RecipTo varchar(4000),
				@TempText nvarchar(4000),
				@CC varchar(4000),
				@BCC varchar(4000),
				@Subject varchar(4000),
				@MsgText varchar(8000),
				@Attachment varchar(4000)

			DECLARE emailqueue_cursor
			CURSOR LOCAL FAST_FORWARD FOR 
			SELECT QueueID, ASRSysEmailQueue.LinkID, RecordID, ASRSysEmailQueue.ColumnID, ColumnValue,RecordDesc,RecalculateRecordDesc,TableID, DateDue
				FROM ASRSysEmailQueue
				INNER JOIN ASRSysEmailLinks ON ASRSysEmailLinks.LinkID = ASRSysEmailQueue.LinkID
				WHERE DateSent IS Null And datediff(dd,DateDue,getdate()) >= 0
				AND (LOWER(@Username) = LOWER([Username]) OR @Username = '''')
				AND TableID is NULL
			UNION SELECT QueueID, LinkID, RecordID, ColumnID, ColumnValue,RecordDesc,RecalculateRecordDesc,TableID, DateDue
					FROM ASRSysEmailQueue
					WHERE DateSent IS Null And datediff(dd,DateDue,getdate()) >= 0
					And (LOWER(@Username) = LOWER([Username]) OR @Username = '''')
					AND TableID is NOT NULL
			ORDER BY dateDue

			OPEN emailqueue_cursor
			FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc, @RecalculateRecordDesc, @TableID, @DateDue

			WHILE (@@fetch_status = 0)
			BEGIN
				IF @RecalculateRecordDesc = 1
					BEGIN	
						IF @ColumnID > 0
							BEGIN
								SELECT @RecDescID = (SELECT RecordDescExprID FROM ASRSYSTables WHERE TableID = 
									(SELECT TableID FROM ASRSysColumns WHERE ColumnID = @ColumnID))
							END
						ELSE IF @TableID > 0
							BEGIN			
								SELECT @RecDescID = (SELECT RecordDescExprID FROM ASRSYSTables WHERE TableID = @TableID)
							END
				
						SET @RecDesc = ''''
						SELECT @sSQL = ''sp_ASRExpr_'' + convert(varchar,@RecDescID)
						IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
						BEGIN
							EXEC @sSQL @RecDesc OUTPUT, @Recordid
						END
					END

				IF @TableID > 0
					BEGIN
						SELECT @TempText = (SELECT TableName FROM ASRSYSTables WHERE TableID = @TableID)
						SET @RecDesc = @TempText + '' : '' + @RecDesc
					END		
			
				IF @ColumnID > 0
					BEGIN
						SELECT @sSQL = ''spASRSysEmailSend_'' + convert(varchar,@LinkID)
						IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
							BEGIN
								SELECT @emailDate = getDate()
								EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, '''', @RecipTo OUTPUT, @CC OUTPUT, @BCC OUTPUT, @Subject OUTPUT, @MsgText OUTPUT, @Attachment OUTPUT
							END
					END
				ELSE IF @TableID > 0
					BEGIN
						SET @sSQL = ''spASRSysEmailAddr''
						IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
							BEGIN
								SELECT @emailDate = getDate()
								EXEC @hResult = @sSQL @RecipTo OUTPUT, @LinkID, 0
								SET @Subject = @columnvalue
								SET @MsgText = @RecDesc
								EXEC spASRSendMail @hResult OUTPUT, @RecipTo, '''', '''', @Subject,  @MsgText, ''''
							END
					END
				IF @hResult = 0
				BEGIN
					UPDATE ASRSysEmailQueue SET DateSent = @emailDate, RepTo = @RecipTo, RepCC = @CC, RepBCC = @BCC, Subject = @Subject, Attachment = @Attachment
					WHERE QueueID = @QueueID
					
					UPDATE ASRSysEmailQueue SET MsgText = @MsgText
					WHERE QueueID = @QueueID
				END
				FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc, @RecalculateRecordDesc, @TableID, @DateDue
			END
			CLOSE emailqueue_cursor
			DEALLOCATE emailqueue_cursor
		END'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 4 of 13 - Optimising Accord Structures'

	IF EXISTS (select * from dbo.sysindexes where id = object_id(N'[dbo].[ASRSysAccordTransactionData]') and Name = 'TransactionID')
		SELECT @NVarCommand = 'CREATE CLUSTERED INDEX [TransactionID] ON [dbo].[ASRSysAccordTransactionData] ([TransactionID], [FieldID]) WITH DROP_EXISTING ON [PRIMARY]'
	ELSE
		SELECT @NVarCommand = 'CREATE CLUSTERED INDEX [TransactionID] ON [dbo].[ASRSysAccordTransactionData] ([TransactionID], [FieldID]) ON [PRIMARY]'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 5 of 13 - Accord Security Setup'

	UPDATE ASRSysPermissionItems SET Description = 'View Transfer Archive' WHERE ItemID = 147


/* ------------------------------------------------------------- */
PRINT 'Step 6 of 13 - Adding Self-service Intranet Views Table'

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysSSIViews]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysSSIViews] (
						[ViewID] [int] NOT NULL ,
						[ButtonLinkPromptText] [varchar] (200) NULL ,
						[ButtonLinkButtonText] [varchar] (200) NULL ,
						[HypertextLinkText] [varchar] (200) NULL ,
						[DropdownListLinkText] [varchar] (200) NULL ,
						[ButtonLink] [bit] NOT NULL ,
						[HypertextLink] [bit] NOT NULL ,
						[DropdownListLink] [bit] NOT NULL  ,
						[SingleRecordView] [bit] NOT NULL ,
						[Sequence] [int] NOT NULL ,
						[LinksLinkText] [varchar] (200) NULL
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand

		SET @sTemp = ''
		SELECT @sTemp = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_SSINTRANET'
				AND parameterKey = 'Param_SelfServiceView'

		IF LEN(@sTemp) > 0
		BEGIN
			SET @iTemp = convert(integer, @sTemp)
			IF @iTemp > 0 
			BEGIN
				SELECT @NVarCommand = 'INSERT INTO ASRSysSSIViews (
						ViewID, 
						ButtonLinkPromptText, 
						ButtonLinkButtonText, 
						HypertextLinkText, 
						DropdownListLinkText, 
						ButtonLink, 
						HypertextLink, 
						DropdownListLink,
						SingleRecordView,
						Sequence,
						LinksLinkText) 
					VALUES (
						' + @sTemp + ',' +
						''''',' +
						''''',' +
						''''',' +
						''''',' +
						'0,' +
						'0,' +
						'0,' +
						'1,' +
						'0,' +
						'''Employee Links'')'
				EXEC sp_executesql @NVarCommand	
			END
		END

		SET @sTemp = ''
		SELECT @sTemp = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_SSINTRANET'
				AND parameterKey = 'Param_LineManagerView'

		IF LEN(@sTemp) > 0
		BEGIN
			SET @iTemp = convert(integer, @sTemp)
			IF @iTemp > 0 
			BEGIN
				SELECT @NVarCommand = 'INSERT INTO ASRSysSSIViews (
						ViewID, 
						ButtonLinkPromptText, 
						ButtonLinkButtonText, 
						HypertextLinkText, 
						DropdownListLinkText, 
						ButtonLink, 
						HypertextLink, 
						DropdownListLink,
						SingleRecordView,
						Sequence,
						LinksLinkText) 
					VALUES (
						' + @sTemp + ',' +
						'''Select ...'',' +
						'''Staff Member'',' +
						'''Select Staff Member'',' +
						''''',' +
						'1,' +
						'1,' +
						'0,' +
						'0,' +
						'1,' +
						'''Staff Member Links'')'
				EXEC sp_executesql @NVarCommand	
			END
		END
	END

/* ------------------------------------------------------------- */
PRINT 'Step 7 of 13 - Amending Self-service Intranet Links Table'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
	and name = 'viewID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD [viewID] [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @sTemp = ''
		SELECT @sTemp = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_SSINTRANET'
				AND parameterKey = 'Param_SelfServiceView'

		IF LEN(@sTemp) > 0
		BEGIN
			SET @iTemp = convert(integer, @sTemp)
			IF @iTemp > 0 
			BEGIN

				SELECT @NVarCommand = 'UPDATE ASRSysSSIntranetLinks SET viewID = ' + @sTemp +
					' WHERE selfServiceAccess = 1'
				EXEC sp_executesql @NVarCommand
			END
		END

		SET @sTemp = ''
		SELECT @sTemp = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_SSINTRANET'
				AND parameterKey = 'Param_LineManagerView'

		IF LEN(@sTemp) > 0
		BEGIN
			SET @iTemp = convert(integer, @sTemp)
			IF @iTemp > 0 
			BEGIN

				SELECT @NVarCommand = 'UPDATE ASRSysSSIntranetLinks SET viewID = ' + @sTemp +
					' WHERE selfServiceAccess = 0'
				EXEC sp_executesql @NVarCommand
			END
		END

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
		and name = 'selfServiceAccess'
	
		if @iRecCount = 1
		BEGIN
			SET @sTemp = ''

			SELECT @sTemp = name
			FROM sysobjects
			WHERE xtype = 'D'
				AND id = (SELECT cdefault FROM syscolumns WHERE id = object_id('ASRSysSSintranetLinks')
				AND name = 'SelfServiceAccess')

			IF LEN(@sTemp) > 0 
			BEGIN
				SET @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks DROP CONSTRAINT [' + @sTemp + ']'
				EXEC sp_executesql @NVarCommand
			END

			if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRIntGetLinkInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
			drop procedure [dbo].[spASRIntGetLinkInfo]
			
			if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRIntGetLinks]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
			drop procedure [dbo].[spASRIntGetLinks]

			SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks DROP COLUMN [selfServiceAccess]'
			EXEC sp_executesql @NVarCommand
		END
	END

	DELETE FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_SSINTRANET'
		AND parameterKey = 'Param_SelfServiceView'

	DELETE FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_SSINTRANET'
		AND parameterKey = 'Param_LineManagerView'

/* ------------------------------------------------------------- */
PRINT 'Step 8 of 13 - Adding Workflow Tables'

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowElementColumns]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowElementColumns] (
						[ID] [int] NOT NULL ,
						[ElementID] [int] NOT NULL ,
						[ColumnID] [int] NOT NULL ,
						[ValueType] [int] NOT NULL ,
						[Value] [varchar] (255) NULL ,
						[WFFormIdentifier] [varchar] (200) NULL ,
						[WFValueIDentifier] [varchar] (200) NULL 
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowElementItems]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowElementItems] (
						[ID] [int] NOT NULL ,
						[ElementID] [int] NOT NULL ,
						[Caption] [varchar] (200) NULL ,
						[DBColumnID] [int] NULL ,
						[DBRecord] [int] NULL ,
						[Identifier] [varchar] (200) NULL ,
						[InputType] [int] NULL ,
						[InputSize] [int] NULL ,
						[InputDecimals] [int] NULL ,
						[InputDefault] [varchar] (200) NULL ,
						[WFFormIdentifier] [varchar] (200) NULL ,
						[WFValueIdentifier] [varchar] (200) NULL ,
						[ItemType] [int] NOT NULL 
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowElements]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowElements] (
						[ID] [int] NOT NULL ,
						[WorkflowID] [int] NOT NULL ,
						[Type] [int] NOT NULL ,
						[Caption] [varchar] (200) NULL ,
						[ConnectionPairID] [int] NULL ,
						[LeftCoord] [int] NOT NULL ,
						[TopCoord] [int] NOT NULL ,
						[DecisionCaptionType] [int] NULL ,
						[Identifier] [varchar] (200) NULL ,
						[TrueFlowIdentifier] [varchar] (200) NULL ,
						[DataAction] [int] NULL ,
						[DataTableID] [int] NULL ,
						[DataRecord] [int] NULL ,
						[EmailID] [int] NULL ,
						[EmailRecord] [int] NULL 
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowInstanceSteps]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowInstanceSteps] (
						[ID] [int] IDENTITY (1, 1) NOT NULL ,
						[InstanceID] [int] NOT NULL ,
						[ElementID] [int] NOT NULL ,
						[Status] [int] NOT NULL 
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowInstanceValues]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowInstanceValues] (
						[ID] [int] NOT NULL ,
						[InstanceID] [int] NOT NULL ,
						[ElementID] [int] NOT NULL ,
						[Identifier] [varchar] (200) NOT NULL ,
						[Value] [varchar] (1000) NULL 
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowInstances]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowInstances] (
						[ID] [int] IDENTITY (1, 1) NOT NULL ,
						[WorkflowID] [int] NOT NULL ,
						[InitiatorID] [int] NOT NULL 
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowLinks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowLinks] (
						[ID] [int] NOT NULL ,
						[WorkflowID] [int] NOT NULL ,
						[StartElementID] [int] NOT NULL ,
						[EndElementID] [int] NOT NULL ,
						[StartOutboundFlowCode] [int] NULL 
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflows]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflows] (
						[id] [int] NOT NULL ,
						[name] [varchar] (255) NOT NULL ,
						[description] [char] (255) NULL ,
						[enabled] [bit] NOT NULL 
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 9 of 13 - Updating intranet permission item descriptions'

	UPDATE ASRSysPermissionItems
	SET Description = 'Data Manager Intranet (Multiple Record access)'
	WHERE ItemID = 4

	UPDATE ASRSysPermissionItems
	SET Description = 'Data Manager Intranet (Single Record access)'
	WHERE ItemID = 100

/* ------------------------------------------------------------- */
PRINT 'Step 10 of 13 - Adding Workflow System Permissions'

	/* Adding System Permissions for Workflow */	
	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 42
	
	SELECT @iRecCount = count(*)
	FROM ASRSysPermissionCategories
	WHERE categoryID = 42

	IF @iRecCount = 0 
	BEGIN

		/* The record doesn't exist, so create it. */
		INSERT INTO ASRSysPermissionCategories
			(categoryID, 
				description, 
				picture, 
				listOrder, 
				categoryKey)
			VALUES(42,
				'Workflow',
				'',
				10,
				'WORKFLOW')

		SELECT @ptrval = TEXTPTR(picture) 
		FROM ASRSysPermissionCategories
		WHERE categoryID = 42

		WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101010000000000028010000160000002800000010000000200000000100040000000000C00000000000000000000000000000000000000000000000000080000080000000808000800000008000800080800000C0C0C000808080000000FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00000000000000000000008000000000000078F87000888888087FFF78008FFFF88FFFFFFF888FFFF8087FFF78008FFFF80078F8700088888800008000000000000000800000000000000080000000000008888888000000008FFFFFFF800000008FFFFFFF800000008FFFFFFF8000000008888888000000000000000000000000FFFF0000F7FF0000C1C0000080C000000000000080C00000C1C00000F7FF0000F7FF0000F7FF000080FF0000007F0000007F0000007F000080FF0000FFFF0000

		DELETE FROM ASRSysPermissionItems WHERE itemid in (150)
		INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
					VALUES (150,'Run',10,42,'RUN')

	END

	-- Give security to admistrators
	SELECT @iRecCount = count(*)
	FROM ASRSysGroupPermissions
	WHERE itemid IN (150)

	IF @iRecCount = 0 
	BEGIN
		INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
			SELECT DISTINCT 150, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
	END


/* ------------------------------------------------------------- */
PRINT 'Step 11 of 13 - Adding new IgnoreZeros column to Custom Reports'

		SELECT @iRecCount = count(id) FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysCustomReportsName')
		and name = 'IgnoreZeros'

		if @iRecCount = 0
		BEGIN
			SELECT @NVarCommand = 'ALTER TABLE [ASRSysCustomReportsName] ADD 
																[IgnoreZeros] [bit] NULL'
			EXEC sp_executesql @NVarCommand
		END

		SELECT @NVarCommand = 'UPDATE [ASRSysCustomReportsName] SET 
																[IgnoreZeros] = 0 WHERE [IgnoreZeros] IS NULL'
		EXEC sp_executesql @NVarCommand


/* ----------------------------------------------------------------- */
PRINT 'Step 12 of 13 - Amending Bradford Factor Merge stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASR_Bradford_MergeAbsences]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASR_Bradford_MergeAbsences]

	SELECT @NVarCommand = 'CREATE PROCEDURE sp_ASR_Bradford_MergeAbsences
	(
	@pdReportStart	  	datetime,
	@pdReportEnd		datetime,
	@pcReportTableName	char(30)
	)
	AS
	BEGIN
		declare @sSql as char(8000)

		/* Variables to hold current absence record */
		declare @pdStartDate as datetime
		declare @pdEndDate as datetime
		declare @pcStartSession as char(2)
		declare @pfDuration as float
		declare @piID as integer
		declare @piPersonnelID as integer
		declare @pbContinuous as bit

		/* Variables to hold last absence record */
		declare @pdLastStartDate as datetime
		declare @pcLastStartSession as char(2)
		declare @pfLastDuration as float
		declare @piLastID as integer
		declare @piLastPersonnelID as integer

		/* Open the passed in table */
		set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, Duration, Absence_ID, Continuous, Personnel_ID FROM '' + @pcReportTableName + '' ORDER BY Personnel_ID, Start_Date ASC''
		execute(@sSQL)
		open BradfordIndexCursor

		/* Loop through the records in the bradford report table */
		Fetch next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID
		while @@FETCH_STATUS = 0
		begin

			if @pbContinuous = 0 Or (@piPersonnelID <> @piLastPersonnelID)
			begin
				Set @pdLastStartDate = @pdStartDate
				Set @pcLastStartSession = @pcStartSession
				Set @pfLastDuration = @pfDuration
				Set @piLastID = @piID

			end
			else
			begin

				Set @pfLastDuration = @pfLastDuration + @pfDuration

				/* update start date */
				set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Start_Date = '''''' + convert(varchar(20),@pdLastStartDate) + '''''', Start_Session = '''''' + @pcLastStartSession + '''''', Duration = '' + Convert(Char(10), @pfLastDuration) + '', Included_Days = '' + Convert(Char(10), @pfLastDuration) + '' Where Absence_ID = '' + Convert(varchar(10),@piId)
				execute(@sSQL)

				/* Delete the previous record from our collection */
				set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Absence_ID = '' + Convert(varchar(10),@piLastId)
				execute(@sSQL)

				Set @piLastID = @piID

			end

			/* Get next absence record */
			Set @piLastPersonnelID = @piPersonnelID
			
			Fetch next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID
		end

		close BradfordIndexCursor
		deallocate BradfordIndexCursor

	END'
	EXEC sp_executesql @NVarCommand



/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step 13 of 13 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '2.19')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '2.19.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v2.19')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v2.19 Of HR Pro'
