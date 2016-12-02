
/* --------------------------------------------------- */
/* Update the database from version 3.7 to version 4.0 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(max),
	@iSQLVersion numeric(3,1),
	@NVarCommand nvarchar(max),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16)

DECLARE @sSQL varchar(max)
DECLARE @sSPCode nvarchar(max)
DECLARE @sSPCode_0 nvarchar(4000)
DECLARE @sSPCode_1 nvarchar(4000)
DECLARE @sSPCode_2 nvarchar(4000)
DECLARE @sSPCode_3 nvarchar(4000)
DECLARE @sSPCode_4 nvarchar(4000)
DECLARE @sSPCode_5 nvarchar(4000)
DECLARE @sSPCode_6 nvarchar(4000)
DECLARE @sSPCode_7 nvarchar(4000)
DECLARE @sSPCode_8 nvarchar(4000)

/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON
SET @DBName = DB_NAME()

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not previous or current version . */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '3.7') and (@sDBVersion <> '3.8') and (@sDBVersion <> '4.0')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on or above SQL2005
SELECT @iSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY('ProductVersion')));
IF (@iSQLVersion < 9)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of HR Pro', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
/* ------------------------------------------------------------- */
PRINT 'Step 1 - Updating Email Definitions'

    DECLARE @MaxLink as nvarchar(max)


	--Add columns to ASRSysEmailLinks
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND name = 'Type')
    BEGIN

		---------------------------------------------------
		--Clear out previous version of v4.0 update script
		IF NOT OBJECT_ID('ASRSysLinkContent', N'U') IS NULL	
			EXEC sp_executesql N'DROP TABLE ASRSysLinkContent'

		IF EXISTS(SELECT id FROM syscolumns
                  WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND name = 'SubjectContentID')
			EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN SubjectContentID'

		IF EXISTS(SELECT id FROM syscolumns
                  WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND name = 'BodyContentID')
			EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN BodyContentID'
		---------------------------------------------------


		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks
			  ADD Type int NULL
                , TableID int NULL
                , DateColumnID int NULL
                , DateOffset int NULL
                , DatePeriod int NULL
                , RecordInsert bit NULL
                , RecordUpdate bit NULL
                , RecordDelete bit NULL
                , SubjectContentID int NULL
                , BodyContentID int NULL'

		EXEC sp_executesql N'UPDATE ASRSysEmailLinks
		      SET Type = CASE WHEN Immediate = 1 THEN 0 ELSE 2 END
                , TableID = (SELECT ASRSysColumns.tableID
                		  FROM   ASRSysColumns
                		  WHERE  ASRSysColumns.columnID = ASRSysEmailLinks.columnID)
                , DateColumnID = CASE WHEN Immediate = 1 THEN 0 ELSE ColumnID END
                , DateOffset = Offset
                , DatePeriod = Period
                , RecordInsert = 0
                , RecordUpdate = 0
                , RecordDelete = 0'

		SELECT @MaxLink = convert(nvarchar(max),max(linkid)) FROM ASRSysEmailLinks
		SET @NVarCommand = 'INSERT ASRSysEmailLinks
			( LinkID
			, ColumnID
			, Title
			, [Immediate]
			, FilterID
			, EffectiveDate
			, [Subject]
			, IncRecDesc
			, Attachment
			, [Type]
			, TableID
			, DateColumnID
			, DateOffset
			, DatePeriod
			, RecordInsert
			, RecordUpdate
			, RecordDelete )
		SELECT
			( (SELECT COUNT(sub.TableID)
			   FROM   ASRSysTables sub 
			   WHERE  sub.TableID <= ASRSysTables.TableID
				 AND  (EmailInsert > 0 OR EmailDelete > 0)) * 2 )
			+'+@MaxLink+' - 1 AS LinkID, 
			0 AS ColumnID,
			''Record Added'' As Title,
			0 AS [Immediate],
			0 as FilterID,
			''01/01/2000'' as EffectiveDate,
			''Record Added'' AS Subject,
			1 as IncRecDesc,
			'''' as Attachment,
			1 as [Type],
			TableID as ''TableID'',
			0 as DateColumnID,
			0 as DateOffset,
			0 as DatePeriod,
			1 as RecordInsert,
			0 as RecordUpdate,
			0 as RecordDelete
		FROM ASRSysTables
		WHERE EmailInsert > 0
		UNION
		SELECT
			( (SELECT COUNT(sub.TableID)
			   FROM   ASRSysTables sub 
			   WHERE  sub.TableID <= ASRSysTables.TableID
				 AND  (EmailInsert > 0 OR EmailDelete > 0)) * 2 )
			+'+@MaxLink+' AS LinkID, 
			0 AS ColumnID,
			''Record Deleted'' As Title,
			0 AS [Immediate],
			0 as FilterID,
			''01/01/2000'' as EffectiveDate,
			''Record Deleted'' AS Subject,
			1 as IncRecDesc,
			'''' as Attachment,
			1 as [Type],
			TableID as ''TableID'',
			0 as DateColumnID,
			0 as DateOffset,
			0 as DatePeriod,
			0 as RecordInsert,
			0 as RecordUpdate,
			1 as RecordDelete
		FROM ASRSysTables
		WHERE EmailDelete > 0'
		EXEC sp_executesql @NVarCommand


		--Add records to ASRSysEmailLinksRecipients for record related links
		SET @NVarCommand = 'INSERT ASRSysEmailLinksRecipients
			( LinkID
			, RecipientID
			, Mode )
		SELECT
			( (SELECT COUNT(sub.TableID)
			   FROM   ASRSysTables sub 
			   WHERE  sub.TableID <= ASRSysTables.TableID
				 AND  (EmailInsert > 0 OR EmailDelete > 0)) * 2 )
			+'+@MaxLink+' - 1 AS LinkID, 
			EmailInsert AS RecipientID,
			0 as Mode
		FROM ASRSysTables
		WHERE EmailInsert > 0
		UNION
		SELECT
			( (SELECT COUNT(sub.TableID)
			   FROM   ASRSysTables sub 
			   WHERE  sub.TableID <= ASRSysTables.TableID
				 AND  (EmailInsert > 0 OR EmailDelete > 0)) * 2 )
			+'+@MaxLink+' AS LinkID, 
			EmailDelete AS RecipientID,
			0 as Mode		FROM ASRSysTables
		WHERE EmailDelete > 0'
		EXEC sp_executesql @NVarCommand

		EXEC sp_executesql N'UPDATE ASRSysEmailLinks
		      SET SubjectContentID = (LinkID*2)-1
                , BodyContentID = (LinkID*2)'


		EXEC sp_executesql N'UPDATE ASRSysEmailQueue
			SET    LinkID =
					(SELECT l.LinkID
					 FROM   ASRSysEmaillinks l
					 WHERE  ASRSysEmailQueue.TableID = l.TableID
						AND l.[Type] = 1
						AND   ((ASRSysEmailQueue.RecalculateRecordDesc = 1 AND l.RecordInsert = 1)
							OR (ASRSysEmailQueue.RecalculateRecordDesc = 0 AND l.RecordDelete = 1)))
			WHERE  ASRSysEmailQueue.ColumnValue IN (''Record Added'',''Record Deleted'')'


		EXEC sp_executesql N'UPDATE ASRSysEmailQueue
			SET    TableID =
					(SELECT l.TableID
					 FROM   ASRSysEmailLinks l
					 WHERE  ASRSysEmailQueue.LinkID = l.LinkID)
			WHERE  ASRSysEmailQueue.TableID is null'


		EXEC sp_executesql N'UPDATE ASRSysEmailQueue
			SET    RepTo =
                    (SELECT Fixed
                     FROM ASRSysEmailAddress
                     JOIN ASRSysTables ON EmailDelete = EmailID
                     WHERE ASRSysTables.TableID = ASRSysEmailQueue.TableID)
				 , RepCC =     ''''
				 , RepBCC =    ''''
				 , [Subject] = ''Record Deleted''
				 , MsgText =
                    (SELECT TableName
                     FROM ASRSysTables
                     WHERE ASRSysTables.TableID = ASRSysEmailQueue.TableID)+'' : ''+RecordDesc
				 , Attachment = ''''
			WHERE  DateSent IS NULL
			   AND ASRSysEmailQueue.ColumnValue = ''Record Deleted'''

	END


	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND name = 'DateAmendment')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks ADD DateAmendment bit NULL'
		EXEC sp_executesql N'UPDATE ASRSysEmailLinks SET DateAmendment = 1'
	END


	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysEmailQueue', 'U') AND name = 'Type')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailQueue ADD Type int NULL'
	END


	--Create table ASRSysEmailLinksColumns FROM ASRSysEmailLinks for column related emails
	IF OBJECT_ID('ASRSysEmailLinksColumns', N'U') IS NULL	
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysEmailLinksColumns]
                    ( [LinkID] [int] NULL
                    , [ColumnID] [int] NULL )
               ON [PRIMARY]'

		EXEC sp_executesql N'INSERT ASRSysEmailLinksColumns
                    ( LinkID
                    , ColumnID )
               SELECT LinkID
                    , ColumnID
               FROM   ASRSysEmailLinks
               WHERE  [Immediate] = 1'
	END


	--Create table ASRSysLinkContent FROM ASRSysEmailLinks
	IF OBJECT_ID('ASRSysLinkContent', N'U') IS NULL	
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysLinkContent]
		       ( [ID] [int] NULL
		       , [ContentID] [int] NULL
		       , [Sequence] [int] NULL
		       , [FixedText] [varchar](max) NULL
		       , [FieldCode] [varchar](1) NULL
		       , [FieldID] [int] NULL )
		       ON [PRIMARY]'

		EXEC sp_executesql N'INSERT ASRSysLinkContent
				   ( ID
				   , ContentID
				   , Sequence
				   , FixedText
				   , FieldCode
				   , FieldID )
			 SELECT  (LinkID*4)-3 as ID
				   , (LinkID*2)-1 as ContentID
				   , 0 as Sequence
				   , l.[Subject] as FixedText
				   , '''' as FieldCode
				   , 0 as FieldID
			 FROM    ASRSysEmailLinks l
			UNION
			 SELECT  (LinkID*4)-2 as ID
				   , (LinkID*2) as ContentID
				   , 0 as Sequence
				   , replace(t.tableName,''_'','' '')+'' : '' as FixedText
				   , ''E'' as FieldCode
				   , t.recordDescExprID as FieldID
			 FROM    ASRSysEmailLinks l
			 JOIN    ASRSysTables t
				  ON t.tableID = l.tableID
			 WHERE   incRecDesc = 1
			     AND t.recordDescExprID > 0
			UNION
			 SELECT  (LinkID*4)-1 as ID
				   , (LinkID*2) as ContentID
				   , 1 as Sequence
				   , CASE WHEN incRecDesc=1 THEN char(13)+char(10) ELSE '''' END
					 +replace(c.columnName,''_'','' '')+'' : '' as FixedText
				   , ''C'' as FieldCode
				   , c.columnID as FieldID
			 FROM    ASRSysEmailLinks l
			 JOIN    ASRSysColumns c
				  ON c.columnID = l.columnID
			 WHERE   incColDetail = 1
			UNION
			 SELECT  (LinkID*4) as ID
				   , (LinkID*2) as ContentID
				   , 2 as Sequence
				   , CASE WHEN incColDetail=1 THEN char(13)+char(10)+char(13)+char(10) ELSE CASE WHEN incRecDesc=1 THEN char(13)+char(10) ELSE '''' END END
					 +char(13)+char(10)+ l.body
				   , '''' as FieldCode
				   , 0 as FieldID
			 FROM    ASRSysEmailLinks l
			 WHERE   IncUserName = 0 and l.body <> ''''
			UNION
			 SELECT  (LinkID*4) as ID
				   , (LinkID*2) as ContentID
				   , 2 as Sequence
				   , CASE WHEN incColDetail=1 THEN char(13)+char(10)+char(13)+char(10) ELSE CASE WHEN incRecDesc=1 THEN char(13)+char(10) ELSE '''' END END
					 +char(13)+char(10)+l.body
					 +char(13)+char(10)+char(13)+char(10)+''Changed By : '' as FixedText
				   , ''X'' as FieldCode
				   , 1 as FieldID
			 FROM    ASRSysEmailLinks l
			 WHERE   IncUserName = 1
			ORDER BY ID'
	END


	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'ColumnID')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN ColumnID'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'Immediate')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN [Immediate]'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'Offset')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN Offset'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'Period')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN Period'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'Subject')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN [Subject]'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'IncRecDesc')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN IncRecDesc'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'IncColDetail')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN IncColDetail'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'IncUserName')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN IncUserName'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'Body')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN Body'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'EmailInsert')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN EmailInsert'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'EmailUpdate')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN EmailUpdate'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysEmailLinks', 'U') AND  name = 'EmailDelete')
		EXEC sp_executesql N'ALTER TABLE ASRSysEmailLinks DROP COLUMN EmailDelete'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysTables', 'U') AND  name = 'EmailInsert')
		EXEC sp_executesql N'ALTER TABLE ASRSysTables DROP COLUMN EmailInsert'

	IF EXISTS(SELECT id FROM syscolumns
	          WHERE  id = OBJECT_ID('ASRSysTables', 'U') AND  name = 'EmailDelete')
		EXEC sp_executesql N'ALTER TABLE ASRSysTables DROP COLUMN EmailDelete'


/* ------------------------------------------------------------- */
PRINT 'Step 2 - Updating Email Procedures'

	DECLARE	@sObjectName varchar(max)
						
	DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
	SELECT name FROM sysobjects
	WHERE name like 'spASRSysEmailSend[_]%' or name like 'spASREmailContent[_]%'

	OPEN tempObjects
	FETCH NEXT FROM tempObjects INTO @sObjectName
	WHILE (@@fetch_status <> -1)
	BEGIN
		EXEC ('DROP PROCEDURE dbo.[' + @sObjectName + ']')
		FETCH NEXT FROM tempObjects INTO @sObjectName
	END
	CLOSE tempObjects
	DEALLOCATE tempObjects


	IF NOT OBJECT_ID('spASREmailImmediate', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[spASREmailImmediate]'

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASREmailImmediate](@Username varchar(255)) AS
	BEGIN
		DECLARE @QueueID int,
				@LinkID int,
				@RecordID int,
				@sSQL nvarchar(max),
				@DateDue datetime,
				@hResult int,
				@TableID int,
				@tmpUser varchar(255),
				@RecalculateRecordDesc int,
				@AttachmentFolder varchar(max)

		DECLARE @TempText nvarchar(max),
				@To varchar(max),
				@CC varchar(max),
				@BCC varchar(max),
				@Subject varchar(max),
				@MsgText varchar(max),
				@Attachment varchar(max)

		SET @AttachmentFolder = ''''
		SELECT @AttachmentFolder = settingvalue
		FROM asrsyssystemsettings
		WHERE [section] = ''email'' and [settingkey] = ''attachment path''

		DECLARE emailqueue_cursor
		CURSOR LOCAL FAST_FORWARD FOR 
		  SELECT ASRSysEmailQueue.QueueID
			   , ASRSysEmailQueue.LinkID
			   , ASRSysEmailQueue.RecordID
			   , ASRSysEmailQueue.TableID
			   , ASRSysEmailQueue.DateDue
			   , ASRSysEmailQueue.UserName
			   , ASRSysEmailQueue.RecalculateRecordDesc
			   , ASRSysEmailQueue.RepTo
               , ASRSysEmailQueue.RepCC
               , ASRSysEmailQueue.RepBCC
               , ASRSysEmailQueue.[Subject]
               , ASRSysEmailQueue.MsgText
               , ASRSysEmailQueue.Attachment
          FROM ASRSysEmailQueue
		  LEFT OUTER JOIN
				 ASRSysEmailLinks
			  ON ASRSysEmailLinks.LinkID = ASRSysEmailQueue.LinkID
		  WHERE  DateSent IS Null
			AND  datediff(dd,DateDue,getdate()) >= 0
			AND  (LOWER(substring(@Username,charindex(''\'',@Username)+1,999)) = LOWER(substring([Username],charindex(''\'',[Username])+1,999))
				  OR @Username = ''''
				  )
		  ORDER BY dateDue

		OPEN emailqueue_cursor
		FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @TableID, @DateDue, @tmpUser, @RecalculateRecordDesc, @To, @CC, @BCC, @Subject, @MsgText, @Attachment

		WHILE (@@fetch_status = 0)
		BEGIN

			SET @hResult = 0
			IF @RecalculateRecordDesc = 1 OR rtrim(isnull(@To,'''')) = ''''
			BEGIN
				SELECT @sSQL = ''spASREmail_'' + convert(varchar,@LinkID)
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					EXEC @hResult = @sSQL @queueid, @recordid, @tmpUser, @To OUTPUT, @CC OUTPUT, @BCC OUTPUT, @Subject OUTPUT, @MsgText OUTPUT, @Attachment OUTPUT
				END
			END

			IF @hResult = 0
			BEGIN
				IF @Attachment <> '''' SET @Attachment = @AttachmentFolder+@Attachment
				EXEC spASRSendMail @hResult OUTPUT, @To, @CC, @BCC, @Subject, @MsgText, @Attachment
				IF @hResult = 0
					UPDATE ASRSysEmailQueue
					SET DateSent = getdate(), RecalculateRecordDesc = 0
					WHERE QueueID = @QueueID
			END

			FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @TableID, @DateDue, @tmpUser, @RecalculateRecordDesc, @To, @CC, @BCC, @Subject, @MsgText, @Attachment
		END

		CLOSE emailqueue_cursor
		DEALLOCATE emailqueue_cursor

	END'
	EXEC sp_executesql @NVarCommand;


/* ------------------------------------------------------------- */
PRINT 'Step 3 - Updating email queue'

	SELECT @NVarCommand = 'DECLARE @iQueueID int,
		@iRecordID int,
		@iRecordDescID int,
		@sRecordDesc varchar(8000),
		@sSQL varchar(max)

	DECLARE emailQueue_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysEmailQueue.queueID, 
		ASRSysEmailQueue.recordID, 
		ASRSysTables.recordDescExprID
	FROM ASRSysEmailQueue
	INNER JOIN ASRSysEmailLinks ON ASRSysEmailQueue.LinkID = ASRSysEmailLinks.LinkID
	INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysEmailLinks.TableID
	WHERE ASRSysEmailQueue.RecordDesc IS NULL
	
	OPEN emailQueue_cursor
	FETCH NEXT FROM emailQueue_cursor INTO @iQueueID, @iRecordID, @iRecordDescID

	WHILE (@@fetch_status = 0)
	BEGIN
		SET @sRecordDesc = ''''
		
		SELECT @sSQL = ''sp_ASRExpr_'' + convert(varchar,@iRecordDescID)
		IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
		BEGIN
			EXEC @sSQL @sRecordDesc OUTPUT, @iRecordID
		END

		UPDATE ASRSysEmailQueue SET RecordDesc = @sRecordDesc WHERE queueid = @iQueueID
		FETCH NEXT FROM emailQueue_cursor INTO @iQueueID, @iRecordID, @iRecordDescID
	END
	CLOSE emailQueue_cursor
	DEALLOCATE emailQueue_cursor'
	EXEC sp_executesql @NVarCommand;
			
			
/* ------------------------------------------------------------- */
PRINT 'Step 4 - Creating Conversion Function'


	IF NOT OBJECT_ID('udfASRConvertNumeric', 'FN') IS NULL	
	  DROP FUNCTION [dbo].[udfASRConvertNumeric]

	SELECT @NVarCommand = 'CREATE FUNCTION [dbo].[udfASRConvertNumeric]
	(
		@in  decimal(15,4)
	  , @dec integer
	  , @sep integer
	)
	RETURNS varchar(MAX)
	AS
	BEGIN

	  DECLARE @out varchar(max)
	  DECLARE @out2 varchar(max)

	  SET @out = convert(varchar(max),cast(@in as money),@sep)
	  SET @out2 = ''''

	  IF @dec <> 2
	  BEGIN
		SET @out = substring(@out,1,CHARINDEX(''.'',@out)-1)
		IF @dec = 1
		  SET @out2 = convert(varchar(max),cast(@in as decimal(15,1)))
		ELSE IF @dec = 3
		  SET @out2 = convert(varchar(max),cast(@in as decimal(15,3)))
		ELSE IF @dec = 4
		  SET @out2 = convert(varchar(max),cast(@in as decimal(15,4)))
		END

		RETURN @out+substring(@out2,CHARINDEX(''.'',@out2),8000)

	END'

	EXEC sp_executesql @NVarCommand;



/* ------------------------------------------------------------- */
PRINT 'Step 5 - Modifying Workflow Data Structures'

	/* ASRSysWorkflowInstanceValues - Add new TempFileUpload_File column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowInstanceValues', 'U')
	AND name = 'TempFileUpload_File'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
							TempFileUpload_File [image] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new TempFileUpload_ContentType column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowInstanceValues', 'U')
	AND name = 'TempFileUpload_ContentType'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
							TempFileUpload_ContentType [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new TempFileUpload_FileName column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowInstanceValues', 'U')
	AND name = 'TempFileUpload_FileName'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
							TempFileUpload_FileName [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 6 - Upgrade system tables to handle new multiline support'

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRResizeColumn]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRResizeColumn]

		SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRResizeColumn]
		(@sTableName	varchar(255)
		,@ColumnName	varchar(255)
		,@Size			varchar(4))
	AS
	BEGIN

		DECLARE @iRecCount integer;
		DECLARE @NVarCommand nvarchar(MAX);

		-- Modify the passed in column to a varchar(max)
		SELECT @iRecCount = COUNT(id) FROM syscolumns
			where [id] = (select id from sysobjects where name = @sTableName)
			and [name] = @ColumnName;

		if @iRecCount = 1
		BEGIN
			SELECT @NVarCommand = ''ALTER TABLE [dbo].['' + @sTableName + ''] '' +
					''ALTER COLUMN ['' + @ColumnName + ''] varchar('' + @Size + '');''
			EXEC sp_executesql @NVarCommand;
		END
	END'
	EXEC sp_executesql @NVarCommand;


	-- Drop constraints on the audit table
	SET @NVarCommand = '';
	SELECT @NVarCommand = @NVarCommand + 'ALTER TABLE ASRSysAuditTrail DROP CONSTRAINT ' 
		+ OBJECT_NAME([constid]) + ';' + CHAR(13)
		FROM sysconstraints
		WHERE OBJECT_NAME([id]) = 'ASRSysAuditTrail'
			AND OBJECT_NAME([constid]) LIKE 'DF%'
	EXEC sp_executesql @NVarCommand;


	-- Upgrade system columns
	EXECUTE spASRResizeColumn 'ASRSysAccordTransactionData','OldData','MAX';
	EXECUTE spASRResizeColumn 'ASRSysAccordTransactionData','NewData','MAX';
	EXECUTE spASRResizeColumn 'ASRSysAccordTransactions','ErrorText','MAX';
	EXECUTE spASRResizeColumn 'ASRSysAccordTransactionWarnings','WarningMessage','MAX';
	EXECUTE spASRResizeColumn 'ASRSysAuditTrail','OldValue','MAX';
	EXECUTE spASRResizeColumn 'ASRSysAuditTrail','NewValue','MAX';
	EXECUTE spASRResizeColumn 'ASRSysColumns','DefaultValue','MAX';
	EXECUTE spASRResizeColumn 'ASRSysColumns','Mask','MAX';
	EXECUTE spASRResizeColumn 'ASRSysControls','Caption','MAX';
	EXECUTE spASRResizeColumn 'ASRSysDataTransferColumns','FromText','MAX';
	EXECUTE spASRResizeColumn 'ASRSysDiaryEvents','EventNotes','MAX';
	EXECUTE spASRResizeColumn 'ASRSysEmailLinks','Body','MAX';
	EXECUTE spASRResizeColumn 'ASRSysEmailQueue','Subject','MAX';
	EXECUTE spASRResizeColumn 'ASRSysEmailQueue','Attachment','MAX';
	EXECUTE spASRResizeColumn 'ASRSysEmailQueue','RepTo','MAX';
	EXECUTE spASRResizeColumn 'ASRSysEmailQueue','RepCC','MAX';
	EXECUTE spASRResizeColumn 'ASRSysEmailQueue','MsgText','MAX';
	EXECUTE spASRResizeColumn 'ASRSysEmailQueue','ColumnValue','MAX';
	EXECUTE spASRResizeColumn 'ASRSysEventLogDetails','Notes','MAX';
	EXECUTE spASRResizeColumn 'ASRSysExportDetails','Data','MAX';	
	EXECUTE spASRResizeColumn 'ASRSysExprComponents','ValueCharacter','MAX';	
	EXECUTE spASRResizeColumn 'ASRSysGlobalItems','Value','MAX';
	EXECUTE spASRResizeColumn 'ASRSysMessages','Message','MAX';
	EXECUTE spASRResizeColumn 'ASRSysOutlookEvents','ErrorMessage','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowElementValidations','Message','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowElementColumns','Value','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowElementItems','Caption','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowElementItems','InputDefault','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceSteps','Message','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceSteps','UserEmail','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceSteps','HypertextLinkedSteps','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceValues','Value','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceValues','ValueDescription','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceValues','TempValue','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceValues','FileUpload_ContentType','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceValues','FileUpload_FileName','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceValues','TempFileUpload_ContentType','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowInstanceValues','TempFileUpload_FileName','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowQueueColumns','ColumnValue','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflows','querystring','MAX';
	EXECUTE spASRResizeColumn 'ASRSysWorkflowStepDelegation','DelegateEmail','MAX';


	SELECT @iRecCount = COUNT(id) FROM syscolumns
		where [id] = (select id from sysobjects where name = 'ASRSysColumns')
		and [name] = 'Size';

	if @iRecCount = 1
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE [dbo].[ASRSysColumns] 
									ALTER COLUMN [Size] integer;'
		EXEC sp_executesql @NVarCommand;
	END


/* ------------------------------------------------------------- */
PRINT 'Step 7 - Column Definition Data'

	/* ASRSysColumns - Remove IsMaxSize column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysColumns', 'U')
	AND name = 'IsMaxSize';

	IF @iRecCount > 0
	BEGIN
		SET @NVarCommand = 'ALTER TABLE [dbo].[ASRSysColumns]
								DROP COLUMN [IsMaxSize];';
		EXEC sp_executesql @NVarCommand;

	END


	-- Update column defintions to maxsize if multiline is selected
	SELECT @iRecCount = COUNT([SettingValue]) FROM [dbo].[ASRSysSystemSettings]
		WHERE [Section] = 'upgrade' AND [SettingKey] = 'varchar(max)'

	IF @iRecCount = 0
	BEGIN
	
		SET @NVarCommand = 'UPDATE [dbo].[ASRSysColumns] SET [size] = 2147483646 WHERE [multiline] = 1';
		EXEC sp_executesql @NVarCommand;

		SET @NVarCommand = 'UPDATE [dbo].[ASRSysCustomReportsDetails] SET [size] = 2147483646 WHERE [size] = 8000';
		EXEC sp_executesql @NVarCommand;

		SET @NVarCommand = 'UPDATE [dbo].[ASRSysExportDetails] SET [FillerLength] = 2147483646 WHERE [FillerLength] = 8000';
		EXEC sp_executesql @NVarCommand;

		SET @NVarCommand = 'UPDATE [dbo].[ASRSysMailMergeColumns] SET [size] = 2147483646 WHERE [size] = 8000';
		EXEC sp_executesql @NVarCommand;

		SET @NVarCommand = 'UPDATE [dbo].[ASRSysMatchReportBreakdown] SET [colsize] = 2147483646 WHERE [colsize] = 8000';
		EXEC sp_executesql @NVarCommand;

		SET @NVarCommand = 'UPDATE [dbo].[ASRSysMatchReportDetails] SET [colsize] = 2147483646 WHERE [colsize] = 8000';
		EXEC sp_executesql @NVarCommand;

		SET @NVarCommand = 'UPDATE [dbo].[ASRSysRecordProfileDetails] SET [size] = 2147483646 WHERE [size] = 8000';
		EXEC sp_executesql @NVarCommand;

		SET @NVarCommand = 'INSERT [dbo].[ASRSysSystemSettings] ([Section], [SettingKey], [SettingValue]) VALUES (''upgrade'', ''varchar(max)'', 1);';
		EXEC sp_executesql @NVarCommand;

	END


/* ------------------------------------------------------------- */
PRINT 'Step 8 - Shared Table Integration Enhancements'

	SET @sSPCode_0 = 'UPDATE ASRSysAccordTransferTypes SET IsVisible = 1 WHERE TransferTypeID IN (5,6,7,8);'
	EXEC sp_executesql @sSPCode_0;


/* ------------------------------------------------------------- */
PRINT 'Step 9 - Removal of unused procedures (Resourcesafication)'

	-- sp_ASRColumnDefault
	IF NOT OBJECT_ID('sp_ASRColumnDefault', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRColumnDefault]'

	-- sp_ASRDropColumn
	IF NOT OBJECT_ID('sp_ASRDropColumn', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRDropColumn]'

	-- sp_ASRDropColumnDefault
	IF NOT OBJECT_ID('sp_ASRDropColumnDefault', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRDropColumnDefault]'

	-- sp_ASRGetChildTables
	IF NOT OBJECT_ID('sp_ASRGetChildTables', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetChildTables]'

	-- sp_ASRGetColumnPrivileges
	IF NOT OBJECT_ID('sp_ASRGetColumnPrivileges', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetColumnPrivileges]'

	-- sp_ASRGetDefaults
	IF NOT OBJECT_ID('sp_ASRGetDefaults', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetDefaults]'

	-- sp_ASRGetGlobalAddDetails
	IF NOT OBJECT_ID('sp_ASRGetGlobalAddDetails', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetGlobalAddDetails]'

	-- sp_ASRGetGlobalUpdateDetails
	IF NOT OBJECT_ID('sp_ASRGetGlobalUpdateDetails', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetGlobalUpdateDetails]'

	-- sp_ASRGetLinkColumns
	IF NOT OBJECT_ID('sp_ASRGetLinkColumns', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetLinkColumns]'

	-- sp_ASRGetLockInfo
	IF NOT OBJECT_ID('sp_ASRGetLockInfo', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetLockInfo]'

	-- sp_ASRGetParentTables
	IF NOT OBJECT_ID('sp_ASRGetParentTables', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetParentTables]'

	-- sp_ASRGetTablePrivileges
	IF NOT OBJECT_ID('sp_ASRGetTablePrivileges', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetTablePrivileges]'

	-- sp_ASRGetViewLinkColumnPrivileges
	IF NOT OBJECT_ID('sp_ASRGetViewLinkColumnPrivileges', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetViewLinkColumnPrivileges]'

	-- sp_ASRGetViewLinkColumns
	IF NOT OBJECT_ID('sp_ASRGetViewLinkColumns', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetViewLinkColumns]'

	-- sp_ASRGetViewPrivileges
	IF NOT OBJECT_ID('sp_ASRGetViewPrivileges', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRGetViewPrivileges]'

	-- sp_ASRPrimaryJoinDetails
	IF NOT OBJECT_ID('sp_ASRPrimaryJoinDetails', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRPrimaryJoinDetails]'

	-- sp_ASRRefreshCalculatedColumns
	IF NOT OBJECT_ID('sp_ASRRefreshCalculatedColumns', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRRefreshCalculatedColumns]'

	-- sp_ASRSetLock
	IF NOT OBJECT_ID('sp_ASRSetLock', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRSetLock]'
	
	-- sp_ASRUserInfo
	IF NOT OBJECT_ID('sp_ASRUserInfo', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[sp_ASRUserInfo]'

	-- spASRSysTableOLEStats
	IF NOT OBJECT_ID('spASRSysTableOLEStats', N'P') IS NULL
		EXEC sp_executesql N'DROP PROCEDURE [dbo].[spASRSysTableOLEStats]'




/* ------------------------------------------------------------- */
PRINT 'Step 10 - Multiline Character Modifications'

	----------------------------------------------------------------------
	-- sp_ASR_AbsenceBreakdown_Run
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASR_AbsenceBreakdown_Run]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASR_AbsenceBreakdown_Run];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASR_AbsenceBreakdown_Run]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASR_AbsenceBreakdown_Run]
		(
			@pdReportStart      datetime,
			@pdReportEnd		datetime,
			@pcReportTableName  char(30)
		) 
		AS 
		BEGIN
			declare @pdStartDate as datetime
			declare @pdEndDate as datetime
			declare @pcStartSession as char(2)
			declare @pcEndSession as char(2)
			declare @pcType as char(50)
			declare @pcRecordDescription as char(100)
		
			declare @pfDuration as float
			declare @pdblSun as float
			declare @pdblMon as float
			declare @pdblTue as float
			declare @pdblWed as float
			declare @pdblThu as float
			declare @pdblFri as float
			declare @pdblSat as float
		
			declare @sSQL as varchar(MAX)
			declare @piParentID as integer
			declare @piID as integer
			declare @pbProcessed as bit
		
			declare @pdTempStartDate as datetime
			declare @pdTempEndDate as datetime
			declare @pcTempStartSession as char(2)
			declare @pcTempEndSession as char(2)
			declare @sTempEndDate as varchar(50)
		
			declare @pfCount as float
			declare @psVer as char(80)
		
			/* Alter the structure of the temporary table so it can hold the text for the days */
			Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ALTER COLUMN Hor NVARCHAR(10)''
			execute(@sSQL)
			Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ADD Processed BIT''
			execute(@sSQL)
			Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ADD DisplayOrder INT''
			execute(@sSQL)
			Set @sSQL = ''ALTER TABLE '' + @pcReportTableName + '' ALTER COLUMN Value decimal(10,5)''
			execute(@sSQL)
		
			/* Load the values from the temporary cursor */
			Set @sSQL = ''DECLARE AbsenceBreakdownCursor CURSOR STATIC FOR SELECT ID, Personnel_ID, Start_Date, End_Date, Start_Session, End_Session, Ver, RecDesc, Processed FROM '' + @pcReportTableName
			execute(@sSQL)
			open AbsenceBreakdownCursor
		
			/* Loop through the records in the absence breakdown report table */
			Fetch Next From AbsenceBreakdownCursor Into @piID, @piParentID, @pdStartDate, @pdEndDate, @pcStartSession, @pcEndSession, @pcType, @pcRecordDescription, @pbProcessed
			while @@FETCH_STATUS = 0
				begin
		
				Set @pdblSun = 0
				Set @pdblMon = 0
				Set @pdblTue = 0
				Set @pdblWed = 0
				Set @pdblThu = 0
				Set @pdblFri = 0
				Set @pdblSat = 0
		
				/* The absence should only calculate for absence within the reporting period */
				set @pdTempStartDate = @pdStartDate
				set @pcTempStartSession = @pcStartSession
				set @pdTempEndDate = @pdEndDate
				set @pcTempEndSession = @pcEndSession
		
				--/* If blank leaving date set it to todays date */
				if @pdTempEndDate is Null set @pdTempEndDate = getdate()
		
				if @pdStartDate <  @pdReportStart
					begin
					set @pdTempStartDate = @pdReportStart
					set @pcTempStartSession = ''AM''
					end
				if @pdTempEndDate >  @pdReportEnd
					begin
					set @pdTempEndDate = @pdReportEnd
					set @pcTempEndSession = ''PM''
					end
		
				set @sTempEndDate = case when @pdEndDate is null then ''null'' else '''''''' + convert(varchar(40),@pdEndDate) + '''''''' end
		
				/* Calculate the days this absence takes up */
				execute sp_ASR_AbsenceBreakdown_Calculate @pfDuration OUTPUT, @pdblMon OUTPUT, @pdblTue OUTPUT, @pdblWed OUTPUT, @pdblThu OUTPUT, @pdblFri OUTPUT, @pdblSat OUTPUT, @pdblSun OUTPUT, @pdTempStartDate, @pcTempStartSession, @pdTempEndDate, @pcTempEndSession, @piParentID
		
				/* Strip out dodgy characters */
				set @pcRecordDescription = replace(@pcRecordDescription,'''''''','''')
				set @pcType = replace(@pcType,'''''''','''')
		
				/* Add Mondays records */
				if @pdblMon > 0
					begin
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 0) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblMon) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',1,1,'' + @sTempEndDate + '',1)''
					execute(@sSQL)
					end
		
				/* Add Tuesday records */
				if @pdblTue > 0
					begin
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 1) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblTue) +  '','''''' + convert(varchar(20),@pdStartDate) + '''''',2,1,'' + @sTempEndDate +'',2)''
					execute(@sSQL)
					end
		
				/* Add Wednesdays records */
				if @pdblWed > 0
					begin
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 2) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblWed) +  '','''''' + convert(varchar(20),@pdStartDate) +  '''''',3,1,'' + @sTempEndDate +'',3)''
					execute(@sSQL)
					end
		
				/* Add new records depending on how many Thursdays were found */
				if @pdblThu > 0
					begin
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 3) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblThu) +  '','''''' + convert(varchar(20),@pdStartDate) + '''''',4,1,'' + @sTempEndDate +'',4)''
					execute(@sSQL)
					end
		
				/* Add new records depending on how many Fridays were found */
				if @pdblFri > 0
					begin
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 4) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblFri) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',5,1,'' + @sTempEndDate +'',5)''
					execute(@sSQL)
					end
		
				/* Add new records depending on how many Saturdays were found */
				if @pdblSat > 0
					begin
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 5) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblSat) + '',''''''+ convert(varchar(20),@pdStartDate) + '''''',6,1,'' + @sTempEndDate +'',6)''
					execute(@sSQL)
					end
		
				/* Add new records depending on how many Sundays were found */
				if @pdblSun > 0
					begin
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '','''''' + DATENAME(weekday, 5) + '''''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pdblSun) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',7,1,'' + @sTempEndDate +'',0)''
					execute(@sSQL)
					end
		
				/* Calculate total duraton of absence */
				set @pfDuration = @pdblMon + @pdblTue + @pdblWed + @pdblThu + @pdblFri + @pdblSat + @pdblSun
		
				if @pfDuration > 0
					begin
					/* Write records for average, totals and count */
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Total'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pfDuration) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',9,1,'' + @sTempEndDate +'',8)''
					execute(@sSQL)
		
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Count'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),1) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',10,1,'' + @sTempEndDate +'',10)''
					execute(@sSQL)
		
					set @sSQL = ''INSERT INTO '' + @pcReportTableName + '' (Personnel_ID, Hor, Ver, RecDesc, Value, Start_Date,Day_Number, Processed, End_Date, DisplayOrder) VALUES ('' + Convert(varchar(10),@piParentID) + '',''''Average'''','''''' + @pcType + '''''', '''''' + @pcRecordDescription + '''''', '' + Convert(varchar(10),@pfDuration) + '','''''' + convert(varchar(20),@pdStartDate) + '''''',9,1,'' + @sTempEndDate +'',9)''
					execute(@sSQL)
					end
		
				/* Process next record */
				Fetch Next From AbsenceBreakdownCursor Into @piID, @piParentID, @pdStartDate, @pdEndDate, @pcStartSession, @pcEndSession, @pcType, @pcRecordDescription, @pbProcessed
		
				end
		
			/* Delete this record from our collection as it''s now been processed */
			set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Processed IS NULL''
			execute(@sSQL)
		
			Set @sSQL = ''DECLARE CalculateAverage CURSOR STATIC FOR SELECT Ver,(SUM(Value) / COUNT(Value)) / COUNT(Value) FROM '' + @pcReportTableName + '' WHERE hor = ''''Average'''' GROUP BY Ver''
			execute(@sSQL)
			open CalculateAverage
		
			Fetch Next From CalculateAverage Into @psVer, @pfCount
			while @@FETCH_STATUS = 0
				begin
		  			Set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Value = '' + Convert(varchar(10),@pfCount) + '' WHERE Ver =  '''''' + @psVer + '''''' AND Hor = ''''Average''''''
				execute(@sSQL)
					Fetch Next From CalculateAverage Into @psVer, @pfCount
				end
		
			/* Tidy up */
			close AbsenceBreakdownCursor
			close CalculateAverage
			deallocate AbsenceBreakdownCursor
			deallocate CalculateAverage
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASR_Bradford_CalculateDurations
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASR_Bradford_CalculateDurations]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASR_Bradford_CalculateDurations];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASR_Bradford_CalculateDurations]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASR_Bradford_CalculateDurations]
		(
			@pdReportStart	  	datetime,
			@pdReportEnd		datetime,
			@pcReportTableName	char(30)
		)
		AS
		BEGIN
		
			declare @pdStartDate as datetime
			declare @pdEndDate as datetime
			declare @pcStartSession as char(2)
			declare @pcEndSession as char(2)
		
			declare @pfDuration as float
			declare @piID as integer
			declare @pfIncludedAmount as float
			declare @sSQL as varchar(MAX)
			declare @pbIncludedRecalculate as bit
			declare @pbTodaysDate as datetime
		
			/* Set an end date in the event of a blank one */
			set @pbTodaysDate = getDate()
		
			/* Open the passed in table */
			set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, End_Date, End_Session, Personnel_ID,Duration FROM '' + @pcReportTableName + '' FOR UPDATE OF Included_Days, Duration''
			execute(@sSQL)
			open BradfordIndexCursor
		
			/* Loop through the records in the bradford report table */
			fetch next from BradfordIndexCursor into @pdStartDate, @pcStartSession, @pdEndDate, @pcEndSession, @piID, @pfDuration
			while @@fetch_status = 0
			begin
				/* Calculate start and end dates */
				Set @pbIncludedRecalculate = 0
		
				/* If empty end date fire off the absence duration calc with system date */
				if isdate(@pdEndDate) = 0
					begin
						execute sp_ASRFn_AbsenceDuration @pfIncludedAmount OUTPUT, @pdStartDate, @pcStartSession, @pbTodaysDate, ''PM'', @piID
						set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Duration = '' + convert(char(10), @pfIncludedAmount) + '' WHERE CURRENT OF BradFordIndexCursor''
						execute(@sSQL)
						set @pdEndDate = @pbTodaysDate
						set @pbIncludedRecalculate = 1
					end
		
				/* Start date is before reporting period */
				if @pdStartDate < @pdReportStart
					begin
						set @pdStartDate = @pdReportStart
						set @pcStartSession = ''AM''
						set @pbIncludedRecalculate = 1
					end
		
				/* End date is outside the reporting period */
				if @pdEndDate > @pdReportEnd
					begin
						set @pdEndDate = @pdReportEnd
						set @pcEndSession = ''PM''
						set @pbIncludedRecalculate = 1
					end
		
				/* If outside of report period, recalculate */
				if @pbIncludedRecalculate = 1
					begin
						execute sp_ASRFn_AbsenceDuration @pfIncludedAmount OUTPUT, @pdStartDate, @pcStartSession, @pdEndDate, @pcEndSession, @piID
						set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Included_Days = '' + convert(char(10), @pfIncludedAmount) + '' WHERE CURRENT OF BradFordIndexCursor''
						execute(@sSQL)
					end
		
				/* Get next record */
				fetch next from BradfordIndexCursor into @pdStartDate, @pcStartSession, @pdEndDate, @pcEndSession, @piID, @pfDuration
			end
		
			close BradfordIndexCursor
			deallocate BradfordIndexCursor
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASR_Bradford_DeleteAbsences
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASR_Bradford_DeleteAbsences]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences]
		(
			@pdReportStart	  	datetime,
			@pdReportEnd		datetime,
			@pbOmitBeforeStart	bit,
			@pbOmitAfterEnd	bit,
			@pcReportTableName	char(30)
		)
		AS
		BEGIN
		
			declare @piID as integer;
			declare @pdStartDate as datetime;
			declare @pdEndDate as datetime;
			declare @iDuration as float;
			declare @pbDeleteThisAbsence as bit;
			declare @sSQL as varchar(MAX);
		
			set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Absence_ID, Start_Date, End_Date, Duration FROM '' + @pcReportTableName;
			execute(@sSQL);
			open BradfordIndexCursor;
		
			Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate, @iDuration;
			while @@FETCH_STATUS = 0
				begin
					set @pbDeleteThisAbsence = 0;
					if @pdEndDate < @pdReportStart set @pbDeleteThisAbsence = 1;
					if @pdStartDate > @pdReportEnd set @pbDeleteThisAbsence = 1;
					if @iDuration = 0 set @pbDeleteThisAbsence = 1;
		
					if @pbOmitBeforeStart = 1 and (@pdStartDate < @pdReportStart)  set @pbDeleteThisAbsence = 1;
					if @pbOmitAfterEnd = 1 and (@pdEndDate > @pdReportEnd)  set @pbDeleteThisAbsence = 1;
		
					if @pbDeleteThisAbsence = 1
						begin
							set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Absence_ID = Convert(Int,'' + Convert(char(10),@piId) + '')'';
							execute(@sSQL);
						end
		
					Fetch Next From BradfordIndexCursor Into @piID, @pdStartDate, @pdEndDate, @iDuration;
				end
		
			close BradfordIndexCursor;
			deallocate BradfordIndexCursor;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASR_Bradford_MergeAbsences
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASR_Bradford_MergeAbsences]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASR_Bradford_MergeAbsences];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASR_Bradford_MergeAbsences]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASR_Bradford_MergeAbsences]
		(
			@pdReportStart	  	datetime,
			@pdReportEnd		datetime,
			@pcReportTableName	char(30)
		)
		AS
		BEGIN
			declare @sSql as varchar(MAX);
		
			/* Variables to hold current absence record */
			declare @pdStartDate as datetime;
			declare @pdEndDate as datetime;
			declare @pcStartSession as char(2);
			declare @pfDuration as float;
			declare @piID as integer;
			declare @piPersonnelID as integer;
			declare @pbContinuous as bit;
		
			/* Variables to hold last absence record */
			declare @pdLastStartDate as datetime;
			declare @pcLastStartSession as char(2);
			declare @pfLastDuration as float;
			declare @piLastID as integer;
			declare @piLastPersonnelID as integer;
		
			/* Open the passed in table */
			set @sSQL = ''DECLARE BradfordIndexCursor CURSOR FOR SELECT Start_Date, Start_Session, Duration, Absence_ID, Continuous, Personnel_ID FROM '' + @pcReportTableName + '' ORDER BY Personnel_ID, Start_Date ASC'';
			execute(@sSQL);
			open BradfordIndexCursor;
		
			/* Loop through the records in the bradford report table */
			Fetch next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID;
			while @@FETCH_STATUS = 0
			begin
		
				if @pbContinuous = 0 Or (@piPersonnelID <> @piLastPersonnelID)
				begin
					Set @pdLastStartDate = @pdStartDate;
					Set @pcLastStartSession = @pcStartSession;
					Set @pfLastDuration = @pfDuration;
					Set @piLastID = @piID;
		
				end
				else
				begin
		
					Set @pfLastDuration = @pfLastDuration + @pfDuration;
		
					/* update start date */
					set @sSQL = ''UPDATE '' + @pcReportTableName + '' SET Start_Date = '''''' + convert(varchar(20),@pdLastStartDate) + '''''', Start_Session = '''''' + @pcLastStartSession + '''''', Duration = '' + Convert(Char(10), @pfLastDuration) + '', Included_Days = '' + Convert(Char(10), @pfLastDuration) + '' Where Absence_ID = '' + Convert(varchar(10),@piId);
					execute(@sSQL);
		
					/* Delete the previous record from our collection */
					set @sSQL = ''DELETE FROM '' + @pcReportTableName + '' Where Absence_ID = '' + Convert(varchar(10),@piLastId);
					execute(@sSQL);
		
					Set @piLastID = @piID;
		
				end
		
				/* Get next absence record */
				Set @piLastPersonnelID = @piPersonnelID;
				
				Fetch next From BradfordIndexCursor Into @pdStartDate, @pcStartSession, @pfDuration, @piID, @pbContinuous, @piPersonnelID;
			end
		
			close BradfordIndexCursor;
			deallocate BradfordIndexCursor;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRAudit
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRAudit]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRAudit]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRAudit]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[sp_ASRAudit] (
			@piColumnID int,
			@piRecordID int,
			@psRecordDesc varchar(255),
			@psOldValue varchar(MAX),
			@psNewValue varchar(MAX))
		AS
		BEGIN
		
			DECLARE @sTableName varchar(128);
			DECLARE @sColumnName varchar(128);
			DECLARE @sUserName varchar(128);
			
			-- Get the table & column name for the given column
			SELECT @sTableName = ASRSysTables.tableName,
				@sColumnName = ASRSysColumns.columnName
			FROM ASRSysColumns
				INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
			WHERE ASRSysColumns.columnID = @piColumnID;
		
			IF @sTableName IS NULL SET @sTableName = ''<Unknown>'';
		
		  SET @sUsername = USER;
			IF UPPER(LEFT(APP_NAME(), 15)) = ''HR PRO WORKFLOW''
				SET @sUsername = ''HR Pro Workflow'';
			ELSE
			BEGIN
				IF EXISTS(SELECT * FROM ASRSysSystemSettings
		                      WHERE [Section] = ''database''
		                        AND [SettingKey] = ''updatingdatedependantcolumns''
		                        AND [SettingValue] = 1)
				BEGIN
				  IF USER = ''dbo''
						SET @sUsername = ''HR Pro Overnight Process'';
				END
			END			
		
			/* Insert a record into the Audit Trail table. */
			INSERT INTO ASRSysAuditTrail 
				(userName, 
				dateTimeStamp, 
				tablename, 
				recordID, 
				recordDesc, 
				columnname, 
				oldValue, 
				newValue,
				ColumnID, 
				deleted)
			VALUES 
				(@sUsername,
				getDate(), 
				@sTableName, 
				@piRecordID, 
				@psRecordDesc, 
				@sColumnName, 
				@psOldValue, 
				@psNewValue,
				@piColumnID,
				0);
				
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- sp_ASRAuditTable
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRAuditTable]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRAuditTable];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRAuditTable]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRAuditTable] (
			@piTableID int,
			@piRecordID int,
			@psRecordDesc varchar(255),
			@psValue varchar(MAX))
		AS
		BEGIN	
			DECLARE @sTableName varchar(128);
		
			/* Get the table name for the given column. */
			SELECT @sTableName = tableName 
				FROM ASRSysTables
				WHERE ASRSysTables.tableID = @piTableID;
		
			IF @sTableName IS NULL SET @sTableName = ''<Unknown>'';
		
			/* Insert a record into the Audit Trail table. */
			INSERT INTO ASRSysAuditTrail 
				(userName, 
				dateTimeStamp, 
				tablename, 
				recordID, 
				recordDesc, 
				columnname, 
				oldValue, 
				newValue,
				ColumnID, 
				Deleted)
			VALUES 
				(CASE
					WHEN UPPER(LEFT(APP_NAME(), 15)) = ''HR PRO WORKFLOW'' THEN ''HR Pro Workflow''
					ELSE user
				END, 
				getDate(), 
				@sTableName, 
				@piRecordID, 
				@psRecordDesc, 
				'''', 
				'''', 
				@psValue,
				0, 
				0);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRCaseSensitiveCompare
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRCaseSensitiveCompare]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRCaseSensitiveCompare];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRCaseSensitiveCompare]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRCaseSensitiveCompare]
		(
			@pfResult		bit OUTPUT,
			@psStringA 		varchar(MAX),
			@psStringB		varchar(MAX)
		)
		AS
		BEGIN
		
			-- Return 1 if the given string are exactly equal.
			DECLARE @iPosition	integer;
		
			SET @pfResult = 0;
		
			IF (@psStringA IS NULL) AND (@psStringB IS NULL) SET @pfResult = 1;
		
			IF (@pfResult = 0) AND (NOT @psStringA IS NULL) AND (NOT @psStringB IS NULL)
			BEGIN
		
				-- LEN() does not look at trailing spaces, so force it too by adding some quotations at the end.
				SET @psStringA = @psStringA + '''''''';
				SET @psStringB = @psStringB + '''''''';
		
				IF LEN(@psStringA) = LEN(@psStringB)
				BEGIN
					SET @pfResult = 1;
		
					SET @iPosition = 1;
					WHILE @iPosition <= LEN(@psStringA) 
					BEGIN
						IF ASCII(SUBSTRING(@psStringA, @iPosition, 1)) <> ASCII(SUBSTRING(@psStringB, @iPosition, 1))
						BEGIN
							SET @pfResult = 0;
							BREAK
						END
		
						SET @iPosition = @iPosition + 1;
					END
				END
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRCrossTabsRecDescs
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRCrossTabsRecDescs]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRCrossTabsRecDescs];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRCrossTabsRecDescs]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRCrossTabsRecDescs]
			(@tablename varchar(8000), @recordDescid int)
		AS
		BEGIN
		
			DECLARE @sSQL nvarchar(MAX);
		
			IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = ''sp_ASRExpr_'' + convert(varchar,@RecordDescID))
			BEGIN
				SET @sSQL = ''
					declare @tableid int;
					declare @recordid int;
					declare @recorddesc varchar(MAX);
		
					DECLARE table_cursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT ID FROM ''+ convert(nvarchar(MAX), @tablename) +''; 
		
					OPEN table_cursor;
					FETCH NEXT FROM table_cursor INTO @recordid;
		
					WHILE (@@fetch_status = 0)
					BEGIN
						exec sp_ASRExpr_'' + convert(nvarchar(128),@RecordDescID) + '' @RecordDesc OUTPUT, @Recordid
						UPDATE '' + convert(nvarchar(128), @tablename) + '' SET RecDesc = @recordDesc WHERE id = @Recordid; 
						FETCH NEXT FROM table_cursor INTO @recordid
					END
					CLOSE table_cursor
					DEALLOCATE table_cursor'';
				EXEC sp_executesql @ssql
		
			END
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRDiaryPurge
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRDiaryPurge]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRDiaryPurge];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRDiaryPurge]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRDiaryPurge]
		AS
		BEGIN
		
			SET NOCOUNT ON;
		
			DECLARE @PurgeDate	datetime,
					@sSQL		nvarchar(MAX),
					@unit		char(1),
		            @period		int,
		            @today		datetime;
		
		    /* Note can''t use sp_ASRPurgeDate as the diary dates include the time !!! */
		
		    select @today = getdate();
		
		    /* Get purge period details */
		    select @unit = unit, @period = (period * -1)
				from asrsyspurgeperiods where purgekey =  ''DIARYSYS'';
		
		    /* calculate purge date */
		    SELECT @purgedate = CASE @unit
		        WHEN ''D'' THEN dateadd(dd,@period,@today)
		        WHEN ''W'' THEN dateadd(ww,@period,@today)
		        WHEN ''M'' THEN dateadd(mm,@period,@today)
		        WHEN ''Y'' THEN dateadd(yy,@period,@today)
		    END;
		
		    SELECT @sSQL = ''DELETE FROM ASRSysDiaryEvents WHERE EventDate < '''''' + 
				convert(varchar,@PurgeDate,101) + '''''' AND ColumnID > 0'';
		
		    EXEC sp_executesql @sSQL;
		
		
		    /* Get purge period details */
		    select @unit = unit, @period = (period * -1)
				from asrsyspurgeperiods where purgekey =  ''DIARYMAN'';
		
		    /* calculate purge date */
		    SELECT @purgedate = CASE @unit
		        WHEN ''D'' THEN dateadd(dd,@period,@today)
		        WHEN ''W'' THEN dateadd(ww,@period,@today)
		        WHEN ''M'' THEN dateadd(mm,@period,@today)
		        WHEN ''Y'' THEN dateadd(yy,@period,@today)
		    END;
		
		    SELECT @sSQL = ''DELETE FROM ASRSysDiaryEvents WHERE EventDate < '''''' 
				+ convert(varchar,@PurgeDate,101) + '' '' + convert(varchar,@PurgeDate,108)
				+ '''''' AND ColumnID = 0'';
		    EXEC sp_executesql @sSQL;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRDropUniqueObject
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRDropUniqueObject]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRDropUniqueObject];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRDropUniqueObject]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRDropUniqueObject](
			@psUniqueObjectName sysname,
			@piType integer)
		AS
		BEGIN
			DECLARE 
				@sCommandString				nvarchar(MAX),
				@sCleanUniqueObjectName		sysname;
		
			/* Clean the input string parameters. */
			SET @sCleanUniqueObjectName = @psUniqueObjectName;
			IF len(@sCleanUniqueObjectName) > 0 SET @sCleanUniqueObjectName = replace(@sCleanUniqueObjectName, '''''''', '''''''''''');
												
			IF (EXISTS (SELECT * 
									FROM sysobjects 
									WHERE name = @psUniqueObjectName))
			BEGIN
				IF @piType = 3 
				BEGIN
					SET @sCommandString = ''DROP TABLE '' + @sCleanUniqueObjectName;
				END
		
				IF @piType = 4
				BEGIN
					SET @sCommandString = ''DROP PROCEDURE '' + @sCleanUniqueObjectName;
				END 
		
				EXECUTE sp_executesql @sCommandString;
		  END
			
			DELETE FROM [dbo].[ASRSysSQLObjects]
			WHERE [Name] = @psUniqueObjectName 
				AND [Type] = @piType
				AND [Owner] = SYSTEM_USER;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_AuditFieldChangedBetweenDates
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_AuditFieldChangedBetweenDates]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_AuditFieldChangedBetweenDates];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_AuditFieldChangedBetweenDates]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_AuditFieldChangedBetweenDates]
		(
			@Result		bit OUTPUT,
			@ColumnID	integer,
			@FromDate	datetime,
			@ToDate		datetime,
			@RecordID	integer
		)
		AS
		BEGIN
			declare @Found as integer;
		
			set @Result = 0;
				
			set @Found = (SELECT Count(DateTimeStamp) FROM [dbo].[ASRSysAuditTrail]
							WHERE ColumnID = @ColumnID
		           				AND RecordID = @RecordID
								AND DateTimeStamp >= @FromDate AND DateTimeStamp <= @ToDate+1);
		
			if @found > 0 set @Result = 1;
		
		End';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_AuditFieldLastChangeDate
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_AuditFieldLastChangeDate]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_AuditFieldLastChangeDate];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_AuditFieldLastChangeDate]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_AuditFieldLastChangeDate]
		(
			@Result		datetime OUTPUT,
			@ColumnID	integer,
			@RecordID	integer
		)
		AS
		BEGIN
			SET @Result = (SELECT TOP 1 DateTimeStamp FROM [dbo].[ASRSysAuditTrail]
					WHERE ColumnID = @ColumnID And @RecordID = RecordID);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_CapitalizeInitials
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_CapitalizeInitials]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_CapitalizeInitials];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_CapitalizeInitials]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_CapitalizeInitials]
		(
			@psResult	varchar(MAX) OUTPUT,
			@psString	varchar(MAX)
		)
		AS
		BEGIN
			DECLARE @iCounter integer;
			DECLARE @sTemp varchar(1);
		
			SET @iCounter = 1;
		
			WHILE @iCounter < LEN(@psString)
			BEGIN
				IF SUBSTRING(@psString, @iCounter, 1) = '' ''
				BEGIN
					SET @sTemp = SUBSTRING(@psString, @iCounter+1, 1);
					SET @psString = STUFF(@psString, @iCounter+1, 1, UPPER(@sTemp));
				END
				ELSE
				BEGIN
					SET @sTemp = SUBSTRING(@psString, @iCounter+1, 1);
					SET @psString = STUFF(@psString, @iCounter+1, 1, LOWER(@sTemp));
				END
		
				SET @iCounter = @iCounter + 1;
			END
		
			-- Change the first letter too
			SET @sTemp = SUBSTRING(@psString, 1, 1);
			SET @psString = STUFF(@psString, 1, 1, UPPER(@sTemp));
		
			SET @psResult = @psString;
			
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_ConvertCharacterToNumeric
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ConvertCharacterToNumeric]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ConvertCharacterToNumeric];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertCharacterToNumeric]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_ConvertCharacterToNumeric]
		(
			@pdblResult			float OUTPUT,
			@psStringToConvert  varchar(MAX)
		)
		AS
		BEGIN
			IF (@psStringToConvert IS NULL) OR (LEN(@psStringToConvert) = 0)
			BEGIN
				SET @pdblResult = 0;
			END
			ELSE
			BEGIN
				IF ISNUMERIC(@psStringToConvert) = 1
				BEGIN
					SET @pdblResult = CONVERT(FLOAT, CONVERT(money, @psStringToConvert));
				END
				ELSE
				BEGIN
					SET @pdblResult = 0;
				END
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_ConvertCurrency
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ConvertCurrency]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ConvertCurrency];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertCurrency]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_ConvertCurrency]
		(
			  @pfResult		float OUTPUT
			, @pfValue		float
			, @psFromCurr	varchar(MAX)
			, @psToCurr		varchar(MAX)
		)
		AS
		BEGIN
		
			DECLARE @sCConvTable 			SysName
					, @sCConvExRateCol		SysName
					, @sCConvCurrDescCol	SysName
					, @sCConvDecCol			SysName
					, @sCommandString		nvarchar(MAX)
					, @sParamDefinition		nvarchar(500);
			
			-- Get the name of the Currency Conversion table and Currency Description column.
			SELECT @sCConvCurrDescCol = ASRSysColumns.ColumnName, @sCConvTable = ASRSysTables.TableName 
			FROM ASRSysModuleSetup 
		   			INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
		            INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID 
			WHERE ASRSysModuleSetup.ModuleKey = ''MODULE_CURRENCY''  AND  ASRSysModuleSetup.ParameterKey = ''Param_CurrencyNameColumn'';
		
			-- Get the name of the Exchange Rate column.
			SELECT @sCConvExRateCol = ASRSysColumns.ColumnName
			FROM ASRSysModuleSetup 
		   			INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
				WHERE ASRSysModuleSetup.ModuleKey = ''MODULE_CURRENCY''  AND  ASRSysModuleSetup.ParameterKey = ''Param_ConversionValueColumn'';
		
			-- Get the name of the Decimals column.
			SELECT @sCConvDecCol = ASRSysColumns.ColumnName
			FROM ASRSysModuleSetup 
		   			INNER JOIN ASRSysColumns ON ASRSysModuleSetup.ParameterValue = ASRSysColumns.ColumnID 
				WHERE ASRSysModuleSetup.ModuleKey = ''MODULE_CURRENCY''  AND  ASRSysModuleSetup.ParameterKey = ''Param_DecimalColumn'';
		
			IF (NOT @sCConvTable IS NULL) AND (NOT @sCConvCurrDescCol IS NULL) AND (NOT @sCConvExRateCol IS NULL) AND (NOT @sCConvDecCol IS NULL) 
		
			-- Create the SQL string that returns the Coverted value.
			BEGIN
				SET @sCommandString = ''SELECT @pfResult = ROUND(ISNULL(('' + LTRIM(RTRIM(STR(@pfValue,20,6))) 
												+ '' / NULLIF((SELECT '' + @sCConvTable + ''.'' + @sCConvExRateCol
											 			  + '' FROM '' + @sCConvTable
														  + '' WHERE '' + @sCConvTable + ''.'' + @sCConvCurrDescCol + '' = '''''' + @psFromCurr + ''''''), 0))''
														  + '' * ''
														  + ''(SELECT '' + @sCConvTable + ''.'' + @sCConvExRateCol
														  + '' FROM '' + @sCConvTable
														  + '' WHERE '' + @sCConvTable + ''.'' + @sCConvCurrDescCol + '' = '''''' + @psToCurr + ''''''), 0)''
														  + '' , ''
														  + '' ISNULL(''
														  + ''(SELECT '' + @sCConvTable + ''.'' + @sCConvDecCol
														  + '' FROM '' + @sCConvTable
														  + '' WHERE '' + @sCConvTable + ''.'' + @sCConvCurrDescCol + '' = '''''' + @psToCurr + ''''''), 0))'';
		
				SET @sParamDefinition = N''@pfResult float output'';
		
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pfResult output;
			END
			ELSE
			BEGIN
				SET @pfResult = NULL;
			END
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_ConvertNumericToString
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ConvertNumericToString]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ConvertNumericToString];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertNumericToString]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_ConvertNumericToString]
		(
			@psResult				varchar(MAX) OUTPUT,
		    @pdblNumericToConvert	float,
		   	@piDecimalPlaces 		integer
		)
		AS
		BEGIN
			/* Convert the number to a string */
			SET @psResult = LTRIM(STR(@pdblNumericToConvert, 20, @piDecimalPlaces));
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_ConvertToLowercase
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ConvertToLowercase]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ConvertToLowercase];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertToLowercase]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE sp_ASRFn_ConvertToLowercase 
		(
			@psResult			varchar(MAX) OUTPUT,
			@psStringToConvert 	varchar(MAX)
		)
		AS
		BEGIN
			SET @psResult = LOWER(@psStringToConvert);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_ConvertToPropercase
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ConvertToPropercase]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ConvertToPropercase];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertToPropercase]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_ConvertToPropercase]
	(
		@psOutput	varchar(MAX) OUTPUT,
		@psInput 	varchar(MAX)
	)
	AS
	BEGIN

		DECLARE @Index	integer,
				@Char	char(1);

		SET @psOutput = LOWER(@psInput);
		SET @Index = 1;
		SET @psOutput = STUFF(@psOutput, 1, 1,UPPER(SUBSTRING(@psInput,1,1)));

		WHILE @Index <= LEN(@psInput)
		BEGIN

			SET @Char = SUBSTRING(@psInput, @Index, 1);

			IF @Char IN (''m'',''M'','' '', '';'', '':'', ''!'', ''?'', '','', ''.'', ''_'', ''-'', ''/'', ''&'','''''''',''('',char(9))
			BEGIN
				IF @Index + 1 <= LEN(@psInput)
				BEGIN
					IF @Char = '''' AND UPPER(SUBSTRING(@psInput, @Index + 1, 1)) != ''S''
						SET @psOutput = STUFF(@psOutput, @Index + 1, 1,UPPER(SUBSTRING(@psInput, @Index + 1, 1)));
					ELSE IF UPPER(@Char) != ''M''
						SET @psOutput = STUFF(@psOutput, @Index + 1, 1,UPPER(SUBSTRING(@psInput, @Index + 1, 1)));

					-- Catch the McName
					IF UPPER(@Char) = ''M'' AND UPPER(SUBSTRING(@psInput, @Index + 1, 1)) = ''C''
					BEGIN
						SET @psOutput = STUFF(@psOutput, @Index + 1, 1,LOWER(SUBSTRING(@psInput, @Index + 1, 1)));
						SET @psOutput = STUFF(@psOutput, @Index + 2, 1,UPPER(SUBSTRING(@psInput, @Index + 2, 1)));
						SET @Index = @Index + 1;
					END
				END
			END

		SET @Index = @Index + 1;
		END

	END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_ConvertToUppercase
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ConvertToUppercase]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ConvertToUppercase];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertToUppercase]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_ConvertToUppercase] 
		(
			@psResult			varchar(MAX) OUTPUT,
			@psStringToConvert	varchar(MAX)
		)
		AS
		BEGIN
			/* Convert the string to upper case */
			SET @psResult = UPPER(@psStringToConvert);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_ExtractCharactersFromTheRight
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ExtractCharactersFromTheRight]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ExtractCharactersFromTheRight];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ExtractCharactersFromTheRight]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_ExtractCharactersFromTheRight]
		(
			@psResult 				varchar(MAX) OUTPUT,
			@psWholeString 			varchar(MAX),
			@piNumberOfCharacters	integer
		)
		AS
		BEGIN
			SET @psResult = RIGHT(@psWholeString, @piNumberOfCharacters);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_ExtractCharactersFromTheLeft
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ExtractCharactersFromTheLeft]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ExtractCharactersFromTheLeft];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ExtractCharactersFromTheLeft]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_ExtractCharactersFromTheLeft]
		(
			@psResult 				varchar(MAX) OUTPUT,
			@psWholeString 			varchar(MAX),
			@piNumberOfCharacters	integer
		)
		AS
		BEGIN
			SET @psResult = LEFT(@psWholeString, @piNumberOfCharacters);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_ExtractPartOfAString
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_ExtractPartOfAString]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_ExtractPartOfAString];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_ExtractPartOfAString]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_ExtractPartOfAString]
		(
			@psResult 				varchar(MAX) OUTPUT,
			@psString 				varchar(MAX),
			@piStart 				integer,
			@piNumberOfCharacters	integer
		)
		AS
		BEGIN
			SET @psResult = SUBSTRING(@psString, @piStart, @piNumberOfCharacters);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_FirstNameFromForenames
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_FirstNameFromForenames]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_FirstNameFromForenames];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_FirstNameFromForenames]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_FirstNameFromForenames]
		(
			@psResult		varchar(MAX) OUTPUT,
			@psForenames	varchar(MAX)
		)
		AS
		BEGIN
			IF (LEN(@psForenames) = 0) OR (@psForenames IS NULL)
			BEGIN
				SET @psResult = '''';
			END
			ELSE
			BEGIN
				IF CHARINDEX('' '', @psForenames) > 0
				BEGIN
					SET @psResult = left(@psForenames, CHARINDEX('' '', @psForenames));
				END
				ELSE
				BEGIN
					SET @psResult = @psForenames;
				END
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_GetCurrentUser
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_GetCurrentUser]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_GetCurrentUser];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_GetCurrentUser]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_GetCurrentUser]
		(
			@psResult	varchar(255) OUTPUT
		)
		AS
		BEGIN
			SET @psResult = 
				CASE 
					WHEN UPPER(LEFT(APP_NAME(), 15)) = ''HR PRO WORKFLOW'' THEN ''HR Pro Workflow'' 
					ELSE SUSER_SNAME()
				END;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_GetFieldFromDatabase
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_GetFieldFromDatabase]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_GetFieldFromDatabase];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_GetFieldFromDatabase]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_GetFieldFromDatabase] (
			@psCharResult		varchar(255) OUTPUT,
			@pfBitResult		bit	OUTPUT,
			@pfltNumResult		float OUTPUT,
			@pdtDateResult		datetime OUTPUT,
			@piSearchColumnID	int,
			@psCharSearchValue	varchar(255),
			@pfBitSearchValue	bit,
			@pfltNumSearchValue	float,
			@pdtDateSearchValue	datetime,
			@piReturnColumnID	int)
		AS
		BEGIN
			DECLARE @sSearchColumnName	sysname,
				@sSearchTableName		sysname,
				@sReturnColumnName		sysname,
				@sReturnTableName		sysname,
				@iSearchColumnType		int,
				@iReturnColumnType		int,
				@sCommandString			nvarchar(MAX),
				@sReturnString			nvarchar(MAX),
				@sParamDefinition		nvarchar(500),
				@sNewCharSearchValue	varchar(MAX),
				@iCharacterIndex		int,
				@iStringLength			int,
				@sCurrentChar 			varchar(1)
		
			/* Replace any single quote characters in the character search string 
			with two single quote characters so that the SQL Select string which is 
			constructed below is still valid for execution. */
			SET @sNewCharSearchValue = '''';
			SET @iCharacterIndex = 0;
			SET @iStringLength = LEN(@psCharSearchValue);
		
			WHILE @iCharacterIndex < @iStringLength
			BEGIN
				SET @iCharacterIndex = @iCharacterIndex + 1;
				SET @sCurrentChar = SUBSTRING(@psCharSearchValue, @iCharacterIndex, 1);
				SET @sNewCharSearchValue = @sNewCharSearchValue + @sCurrentChar;
			
				IF @sCurrentChar = ''''''''
				BEGIN
					SET @sNewCharSearchValue = @sNewCharSearchValue + @sCurrentChar;
				END
			END
		
			SET @psCharSearchValue = @sNewCharSearchValue;
		
			/* Get the name of the search column. */
			SELECT @sSearchColumnName = ASRSysColumns.columnName, 
				@sSearchTableName = ASRSysTables.tableName, 
				@iSearchColumnType = ASRSysColumns.dataType
			FROM ASRSysColumns
			JOIN ASRSysTables 
				ON ASRSysTables.tableID = ASRSysColumns.tableID
			WHERE ASRSysColumns.columnID = @piSearchColumnID;
		
			/* Get the name of the return column. */
			SELECT @sReturnColumnName = ASRSysColumns.columnName, 
				@sReturnTableName = ASRSysTables.tableName, 
				@iReturnColumnType = ASRSysColumns.dataType
			FROM ASRSysColumns
			JOIN ASRSysTables 
				ON ASRSysTables.tableID = ASRSysColumns.tableID
			WHERE ASRSysColumns.columnID = @piReturnColumnID;
		
			IF (NOT @sSearchColumnName IS NULL) 
				AND (NOT @sSearchTableName IS NULL) 
				AND(NOT @sReturnColumnName IS NULL) 
				AND (NOT @sReturnTableName IS NULL)
				AND ((@iSearchColumnType = 12) OR (@iSearchColumnType = -7) OR (@iSearchColumnType = 4) OR (@iSearchColumnType = 2) OR (@iSearchColumnType = 11)) 
				AND ((@iReturnColumnType = 12) OR (@iReturnColumnType = -7) OR (@iReturnColumnType = 4) OR (@iReturnColumnType = 2) OR (@iReturnColumnType = 11)) 
				AND (@sSearchTableName = @sReturnTableName)
			BEGIN
				IF @iReturnColumnType = 12 
				BEGIN
					SET @sReturnString = ''@charResult'';
					SET @sParamDefinition = N''@charResult varchar(255) OUTPUT'';
				END
		
				IF @iReturnColumnType = -7 
				BEGIN
					SET @sReturnString = ''@bitResult'';
					SET @sParamDefinition = N''@bitResult bit OUTPUT'';
				END
		
				IF (@iReturnColumnType = 4) OR (@iReturnColumnType = 2) 
				BEGIN
					SET @sReturnString = ''@numResult'';
					SET @sParamDefinition = N''@numResult float OUTPUT'';
				END
		
				IF @iReturnColumnType = 11 
				BEGIN
					SET @sReturnString = ''@datetimeResult'';
					SET @sParamDefinition = N''@dateResult datetime OUTPUT'';
				END
		
				IF @iSearchColumnType = 12 
				BEGIN
					SET @sCommandString = ''SELECT '' + @sReturnString + '' = '' + @sReturnColumnName + '' FROM '' + @sReturnTableName + '' WHERE '' + @sSearchColumnName + '' = '''''' + @psCharSearchValue + '''''''';
				END
		
				IF @iSearchColumnType = -7 
				BEGIN
					SET @sCommandString = ''SELECT  '' + @sReturnString + '' = '' + @sReturnColumnName + '' FROM '' + @sReturnTableName + '' WHERE '' + @sSearchColumnName + '' = '' + convert(varchar(MAX), @pfBitSearchValue);
				END
		
				IF (@iSearchColumnType = 4) OR (@iSearchColumnType = 2) 
				BEGIN
					SET @sCommandString = ''SELECT  '' + @sReturnString + '' = '' + @sReturnColumnName + '' FROM '' + @sReturnTableName + '' WHERE '' + @sSearchColumnName + '' = '' + convert(varchar(MAX), @pfltNumSearchValue)
				END
			
				IF @iSearchColumnType = 11 
				BEGIN
					SET @sCommandString = ''SELECT  '' + @sReturnString + '' = '' + @sReturnColumnName + '' FROM '' + @sReturnTableName + '' WHERE '' + @sSearchColumnName + '' = '''''' + convert(varchar(MAX), @pdtDateSearchValue, 101) + ''''''''
				END
				IF @iReturnColumnType = 12 EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psCharResult OUTPUT;
				IF @iReturnColumnType = -7 EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pfBitResult OUTPUT;
				IF (@iReturnColumnType = 4) OR (@iReturnColumnType = 2) EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pfltNumResult OUTPUT;
				IF @iReturnColumnType = 11 EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pdtDateResult OUTPUT;
			END
		
			/* Return the result. */
			IF @iReturnColumnType = 12 SELECT @psCharResult AS result;
			IF @iReturnColumnType = -7 SELECT @pfBitResult AS result;
			IF ((@iReturnColumnType = 4) OR (@iReturnColumnType = 2)) SELECT @pfltNumResult AS result;
			IF @iReturnColumnType = 11 SELECT @pdtDateResult AS result;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_GetUniqueCode
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_GetUniqueCode]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_GetUniqueCode];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_GetUniqueCode]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_GetUniqueCode]
		(
			@piResult		int OUTPUT,
			@psCodePrefix	varchar(MAX) = '''',
			@piSuffixRoot	int=1
		)
		AS
		BEGIN
			DECLARE @iOldCodeSuffix int;
			DECLARE @iNewCodeSuffix int;
		
			-- Get the current maximum code suffix for the given code prefix.
			SELECT @iOldCodeSuffix = maxCodeSuffix 
				FROM [dbo].[ASRSysUniqueCodes]
				WHERE codePrefix = @psCodePrefix;
		
			IF @iOldCodeSuffix IS NULL 
			BEGIN
				-- The given code prefix DOES NOT exist in the database, so set the suffix to be the given root suffix, and insert the new code into the database.
				SELECT @iNewCodeSuffix = @piSuffixRoot;
				INSERT INTO [dbo].[ASRSysUniqueCodes] (codePrefix, maxCodeSuffix) VALUES (@psCodePrefix, @iNewCodeSuffix);
			END
			ELSE
			BEGIN
				-- The given code prefix DOES exist in the database, so set the suffix to be the current max suffix plus 1, and update the code into the database.
				SELECT @iNewCodeSuffix = @iOldCodeSuffix + 1;
				UPDATE [dbo].[ASRSysUniqueCodes] SET maxCodeSuffix = @iNewCodeSuffix WHERE codePrefix = @psCodePrefix;
			END
		
			-- Return the new code suffix.
			SET @piResult = @iNewCodeSuffix;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_IfThenElse_3_1_1
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_IfThenElse_3_1_1]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_IfThenElse_3_1_1];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_IfThenElse_3_1_1]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_IfThenElse_3_1_1]
		(
			@psResult   	varchar(MAX) OUTPUT,
			@pfTestValue	bit,
			@psString1		varchar(MAX),
			@psString2		varchar(MAX)
		)
		AS
		BEGIN
			IF @pfTestValue = 1
			BEGIN
				SET @psResult = @psString1;
			END
			ELSE
			BEGIN
				SET @psResult = @psString2;
			END	
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_InitialsFromForenames
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_InitialsFromForenames]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_InitialsFromForenames];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_InitialsFromForenames]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_InitialsFromForenames]
		(
			@psResult 		varchar(MAX) OUTPUT,
			@psForenames	varchar(MAX)
		)
		AS
		BEGIN
			DECLARE @iCounter	integer;
		
			SET @iCounter = 1
		
			IF LEN(@psForenames) > 0 
			BEGIN
				SET @psResult = UPPER(left(@psForenames,1));
		
				WHILE @iCounter < LEN(@psForenames)
				BEGIN
					IF SUBSTRING(@psForenames, @iCounter, 1) = '' ''
					BEGIN
						SET @psResult = @psResult + UPPER(SUBSTRING(@psForenames, @iCounter+1, 1));
					END
			
					SET @iCounter = @iCounter + 1;
				END
		
				SET @psResult = @psResult + '' '';
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_IsEmpty
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_IsEmpty]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_IsEmpty];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_IsEmpty]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_IsEmpty]
		(
		    @result integer OUTPUT,
		    @vartotest varchar(MAX)
		)
		AS
		BEGIN
			
			IF LEN(@vartotest) = 0 
				SET @result = 1;
		
			IF @vartotest IS NULL
				SET @result = 1;
		
			IF LEN(@vartotest) > 0
				SET @result = 0;
		
			SELECT @result AS result;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_IsEmpty_1
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_IsEmpty_1]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_IsEmpty_1];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_IsEmpty_1]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_IsEmpty_1]
		(
			@pfResult	bit OUTPUT,
			@psString	varchar(MAX)
		)
		AS
		BEGIN
			SET @pfResult = 0;
		
			IF LEN(@psString) = 0 
				SET @pfResult = 1;
		
			IF @psString IS null
				SET @pfResult = 1;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_IsOvernightProcess
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_IsOvernightProcess]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_IsOvernightProcess];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_IsOvernightProcess]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_IsOvernightProcess]
		(
		    @result integer OUTPUT
		)
		AS
		BEGIN
		
			DECLARE @iCount					integer;
			DECLARE	@sTempExecString		nvarchar(MAX);
			DECLARE	@sTempParamDefinition	nvarchar(500);
			DECLARE	@sValue					varchar(MAX);
		
			/* Check if the ''ASRSysSystemSettings'' table exists. */
			SELECT @iCount = COUNT(*)
				FROM sysobjects 
				WHERE name = ''ASRSysSystemSettings'';
				
			IF @iCount = 1
			BEGIN
				/* The ASRSysSystemSettings table exists. See if the required records exists in it. */
				SET @sTempExecString = ''SELECT @sValue = settingValue'' +
					'' FROM ASRSysSystemSettings'' +
					'' WHERE section = ''''database'''''' +
					'' AND settingKey = ''''updatingdatedependantcolumns'''''';
				SET @sTempParamDefinition = N''@sValue varchar(MAX) OUTPUT'';
				EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sValue OUTPUT;
			
				IF NOT @sValue IS NULL
				BEGIN
					SET @result = CONVERT(bit, @sValue);
				END
			END
			ELSE
			BEGIN
				SELECT @result = [UpdatingDateDependentColumns] FROM [dbo].[ASRSysConfig];
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_IsPopulated_1
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_IsPopulated_1]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_IsPopulated_1];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_IsPopulated_1]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_IsPopulated_1]
		(
			@pfResult	bit OUTPUT,
			@psString	varchar(MAX)
		)
		AS
		BEGIN
			SET @pfResult = 1;
		
			IF LEN(@psString) = 0 SET @pfResult = 0;
		
			IF @psString IS NULL SET @pfResult = 0;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_LengthOfCharacterField
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_LengthOfCharacterField]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_LengthOfCharacterField];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_LengthOfCharacterField]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_LengthOfCharacterField]
		(
			@piResult 		integer OUTPUT,
			@psWholeString	varchar(MAX)
		)
		AS
		BEGIN
			SET @piResult = LEN(@psWholeString);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_NameOfDay
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_NameOfDay]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_NameOfDay];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_NameOfDay]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_NameOfDay]
		(
			@psResult	varchar(MAX) OUTPUT,
			@pdtDate 	datetime
		)
		AS
		BEGIN
			SET @psResult = DATENAME(weekday, @pdtDate);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_NameOfMonth
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_NameOfMonth]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_NameOfMonth];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_NameOfMonth]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_NameOfMonth]
		(
			@psResult	varchar(MAX) OUTPUT,
			@pdtDate 	datetime
		)
		AS
		BEGIN
			SET @psResult = DATENAME(month, @pdtDate);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_NiceDate
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_NiceDate]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_NiceDate];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_NiceDate]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_NiceDate]
		(
			@psResult	varchar(MAX) OUTPUT,
			@pdtDate 	datetime
		)
		AS
		BEGIN
			
			-- Format(pvParam1, "dddd, mmmm d yyyy")
			IF @pdtDate IS NULL
			BEGIN
				SET @psResult = '''';
			END
			ELSE
			BEGIN
				SET @psResult = datename(dw, @pdtDate) + '', '' + datename(mm, @pdtDate) + '' '' + ltrim(str(datepart(dd, @pdtDate))) + '' '' + ltrim(str(datepart(yy, @pdtDate)));
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_NiceTime
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_NiceTime]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_NiceTime];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_NiceTime]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_NiceTime]
		(
			@psResult 		varchar(MAX) OUTPUT,
			@psTimeString	varchar(MAX) -- in the format hh:mm:ss (24 hour clock)
		)
		AS
		BEGIN
		
			-- Return the given time in the format hh:mm am/pm (12 hour clock)
			select @psResult = 
			case 
				when len(ltrim(rtrim(@psTimeString))) = 0 then ''''
				else 
					case 
						when isdate(@psTimeString) = 0 then ''***''
						else (convert(varchar(2),((datepart(hour,convert(datetime, @psTimeString)) + 11) % 12) + 1)
							+ '':'' + right(''00'' + datename(minute, convert(datetime, @psTimeString)),2)
							+ case 
								when datepart(hour, convert(datetime, @psTimeString)) > 11 then '' pm''
								else '' am'' 
							end) 
					end 
			end
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_NumberOfWorkingDaysPerWeek
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_NumberOfWorkingDaysPerWeek]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_NumberOfWorkingDaysPerWeek];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_NumberOfWorkingDaysPerWeek]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_NumberOfWorkingDaysPerWeek]
		(
			@pdblResult 	float OUTPUT,
			@psPattern 		varchar(MAX)		
			/* Working pattern. 14 characters long in the format ''SsMmTtWwTtFfSs''
			where a uppercase letter relates to the morning, and the lowercase letter relates to the afternnon of the appropriate day. 
			A space means that the morning/afternoon is not worked, anything else means that the session is worked. */
		)
		AS
		BEGIN
			DECLARE @iCounter	integer;
		
			SET @pdblResult = 0;
			SET @iCounter = 0;
		
			WHILE @iCounter <= LEN(@psPattern)
			BEGIN
				IF SUBSTRING(@psPattern, @iCounter, 1) <> '' ''
				BEGIN
					SET @pdblResult = @pdblResult + 0.5;
				END
		
				SET @iCounter = @iCounter + 1;
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_RemoveLeadingAndTrailingSpaces
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_RemoveLeadingAndTrailingSpaces]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_RemoveLeadingAndTrailingSpaces];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_RemoveLeadingAndTrailingSpaces]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_RemoveLeadingAndTrailingSpaces]
		(
			@psResult		varchar(MAX) OUTPUT,
			@psTextToTrim	varchar(MAX)
		)
		AS
		BEGIN
			SET @psResult = LTRIM(RTRIM(@psTextToTrim));
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_SearchForCharacterString
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_SearchForCharacterString]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_SearchForCharacterString];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_SearchForCharacterString]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_SearchForCharacterString]
		(	
			@piResult		integer OUTPUT,
			@psWholeString	varchar(MAX),
			@psSearchString	varchar(MAX)
		)
		AS
		BEGIN
			SET @piResult = CHARINDEX(@psSearchString, @psWholeString);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_StatutorySickPay
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_StatutorySickPay]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_StatutorySickPay];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_StatutorySickPay]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_StatutorySickPay]
		(
			@piAbsenceRecordID		int
		)
		AS
		BEGIN
			/* Refresh the SSP fields in the Absence records for the Personnel record that is the parent of the given Absence record ID. */
		
			/* Absence module - Personnel table variables. */
			DECLARE @iPersonnelTableID				integer,
				@sPersonnelTableName 				varchar(128),
				@sWorkingDaysNum_ColumnName 		varchar(128),
				@sWorkingDaysPattern_ColumnName 	varchar(128),
				@sDateOfBirth_ColumnName 			varchar(128);
		
			/* Personnel record variables. */
			DECLARE @iPersonnelRecordID 			integer,
				@iWorkingDaysPerWeek 				integer,
				@sWorkingPattern 					varchar(MAX),
				@dtDateOfBirth						datetime,
				@dtRetirementDate					datetime,
				@dtSixteenthBirthday				datetime;
		
			/* Absence module - Absence table variables. */
			DECLARE @sAbsenceTableName				varchar(128),
				@sAbsence_StartDateColumnName		varchar(128),
				@sAbsence_EndDateColumnName			varchar(128),
				@sAbsence_StartSessionColumnName	varchar(128),
				@sAbsence_EndSessionColumnName		varchar(128),
				@sAbsence_TypeColumnName			varchar(128),
				@sAbsence_SSPAppliesColumnName 		varchar(128),
				@sAbsence_QualifyingDaysColumnName 	varchar(128),
				@sAbsence_WaitingDaysColumnName 	varchar(128),
				@sAbsence_PaidDaysColumnName 		varchar(128),
				@iAbsence_WorkingDaysType 			integer;
		
			/* Absence record variables. */
			DECLARE @cursAbsenceRecords			cursor,
				@cursFollowingAbsenceRecords	cursor,
				@iAbsenceRecordID 				integer,
				@dtStartDate 					datetime,
				@dtEndDate 						datetime,
				@sStartSession					varchar(MAX),
				@sEndSession 					varchar(MAX),
				@dtWholeStartDate 				datetime,
				@dtWholeEndDate 				datetime,
				@dtFollowingStartDate 			datetime,
				@dtFollowingEndDate 			datetime,
				@sFollowingStartSession	 		varchar(MAX),
				@sFollowingEndSession 			varchar(MAX),
				@dtFollowingWholeStartDate 		datetime,
				@dtFollowingWholeEndDate 		datetime,
				@fOriginalSSPApplies			bit,
				@dblOriginalQualifyingDays		float,
				@dblOriginalWaitingDays			float,
				@dblOriginalPaidDays			float;
		
			/* Absence module - Absence Type table variables. */
			DECLARE @sAbsenceTypeTableName			varchar(128),
				@sAbsenceType_TypeColumnName		varchar(128),
				@sAbsenceType_SSPAppliesColumnName	varchar(128);
		
			/* General procedure handling variables. */
			DECLARE @fOK	 					bit,
				@iLoop							integer,
				@iIndex							integer,
				@sCommandString					nvarchar(MAX),
				@sParamDefinition				nvarchar(500),
				@dblWaitEntitlement 			float,
				@dblAbsenceEntitlement 			float,
				@dblQualifyingDays 				float,
				@dblWaitingDays 				float,
				@dblPaidDays 					float,
				@fSSPApplies					bit,
				@dtTempDate						datetime,
				@fAddOK							bit,
				@dblAddAmount					float,
				@fContinue 						bit,
				@iConsecutiveRecords			integer,
				@dtConsecutiveStartDate 		datetime,
				@dtConsecutiveEndDate 			datetime,
				@dtConsecutiveWholeStartDate 	datetime,
				@dtConsecutiveWholeEndDate 		datetime,
				@sConsecutiveStartSession 		varchar(MAX),
				@sConsecutiveEndSession 		varchar(MAX),
				@dtLastWholeEndDate 			datetime,
				@dtFirstLinkedWholeStartDate 	datetime,
				@iYearDifference 				integer;
		
			SET @fOK = 1;
		
			/* Get the Absence module parameters. */
			/* Get the Personnel table name and ID. */
			SELECT @iPersonnelTableID = convert(integer, parameterValue), 
				@sPersonnelTableName = ASRSysTables.tableName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysTables 
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysTables.tableID
			WHERE moduleKey = ''MODULE_PERSONNEL''
			AND parameterKey = ''Param_TablePersonnel'';
		
			/* Get the Personnel - Date Of Birth column name. */
			SELECT @sDateOfBirth_ColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_PERSONNEL''
			AND parameterKey = ''Param_FieldsDateOfBirth'';
		
			/* Get the Absence table name. */
			SELECT @sAbsenceTableName = ASRSysTables.tableName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysTables 
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysTables.tableID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_TableAbsence'';
		
			/* Get the Absence - Start Date column name. */
			SELECT @sAbsence_StartDateColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldStartDate'';
		
			/* Get the Absence - End Date column name. */
			SELECT @sAbsence_EndDateColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldEndDate'';
		
			/* Get the Absence - Start Session column name. */
			SELECT @sAbsence_StartSessionColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldStartSession'';
		
			/* Get the Absence - End Session column name. */
			SELECT @sAbsence_EndSessionColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldEndSession'';
		
			/* Get the Absence - Type column name. */
			SELECT @sAbsence_TypeColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldType'';
		
			/* Get the Absence - SSP Applies column name. */
			SELECT @sAbsence_SSPAppliesColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldSSPApplies'';
		
			/* Get the Absence - Qualifying Days column name. */
			SELECT @sAbsence_QualifyingDaysColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldQualifyingDays'';
		
			/* Get the Absence - Waiting Days column name. */
			SELECT @sAbsence_WaitingDaysColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldWaitingDays'';
		
			/* Get the Absence - Paid Days column name. */
			SELECT @sAbsence_PaidDaysColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldPaidDays'';
		
			/* Get the Absence - Working Days selection type. */
			SELECT @iAbsence_WorkingDaysType = convert(integer, parameterValue)
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_WorkingDaysType'';
		
			/* Get the Absence Type table name. */
			SELECT @sAbsenceTypeTableName = ASRSysTables.tableName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysTables
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysTables.tableID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_TableAbsenceType'';
		
			/* Get the Absence Type - Type column name. */
			SELECT @sAbsenceType_TypeColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldTypeType'';
		
			/* Get the Absence Type - SSP Applies column name. */
			SELECT @sAbsenceType_SSPAppliesColumnName = ASRSysColumns.columnName
			FROM ASRSysModuleSetup
			INNER JOIN ASRSysColumns
				ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
			WHERE moduleKey = ''MODULE_ABSENCE''
			AND parameterKey = ''Param_FieldTypeSSP'';
		
			/* Validate the Absence module variables. */
			IF (@iPersonnelTableID IS null)
				OR (@sPersonnelTableName IS null)
				OR (@sAbsenceTableName IS null) 
				OR (@sAbsence_StartDateColumnName IS null) 
				OR (@sAbsence_EndDateColumnName IS null) 
				OR (@sAbsence_StartSessionColumnName IS null) 
				OR (@sAbsence_EndSessionColumnName IS null)  
				OR (@sAbsence_TypeColumnName IS null)   
				OR (@sAbsence_SSPAppliesColumnName IS null)
				OR (@sAbsence_QualifyingDaysColumnName IS null)
				OR (@sAbsence_WaitingDaysColumnName IS null)
				OR (@sAbsence_PaidDaysColumnName IS null)
				OR (@iAbsence_WorkingDaysType IS null)
				OR (@sAbsenceTypeTableName IS null)   
				OR (@sAbsenceType_TypeColumnName IS null)    
				OR (@sAbsenceType_SSPAppliesColumnName IS null) SET @fOK = 0;
		
			IF @fOK = 1
			BEGIN
				/* Get the ID  of the associated record in the Personnel table. */
				SET @sParamDefinition = N''@recordID integer OUTPUT'';
				SET @sCommandString = ''SELECT @recordID = id_'' + convert(varchar(128), @iPersonnelTableID) + 
					'' FROM '' + @sAbsenceTableName + 
					'' WHERE id = '' + convert(varchar(128), @piAbsenceRecordID);
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iPersonnelRecordID OUTPUT;
		
				IF (@iPersonnelRecordID IS null) OR (@iPersonnelRecordID <= 0) SET @fOK = 0;
			END
		
			IF (@fOK = 1) AND (NOT @sDateOfBirth_ColumnName IS null) 
			BEGIN
				/* Get the retirement date, and the date of the person''s sixteenth birthday. */
				SET @sParamDefinition = N''@dateOfBirth datetime OUTPUT'';
				SET @sCommandString = ''SELECT @dateOfBirth = convert(datetime, convert(varchar(20), '' + @sDateOfBirth_ColumnName + '', 101))'' +
					'' FROM '' + @sPersonnelTableName + 
					'' WHERE id = '' + convert(varchar(128), @iPersonnelRecordID);
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @dtDateOfBirth OUTPUT;
		
				IF (NOT @dtDateOfBirth IS null) SET @dtRetirementDate = dateadd(yy, 65, @dtDateOfBirth);
				IF (NOT @dtDateOfBirth IS null) SET @dtSixteenthBirthday = dateadd(yy, 16, @dtDateOfBirth);
			END
		
			IF @fOK = 1 
			BEGIN
				/* Get the number of working days per week. */
				SET @iWorkingDaysPerWeek = 0;
				SET @sWorkingPattern = '''';
		
				IF @iAbsence_WorkingDaysType = 0	/* The Working Days are an straight numeric value. */
				BEGIN
					SELECT @iWorkingDaysPerWeek = convert(integer, parameterValue)
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_ABSENCE''
					AND parameterKey = ''Param_WorkingDaysNum'';
		
					IF @iWorkingDaysPerWeek IS null SET @fOK = 0;
				END
		
				IF @iAbsence_WorkingDaysType = 1	/* The Working Days are an straight working pattern value. */
				BEGIN
					SELECT @sWorkingPattern = parameterValue
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_ABSENCE''
					AND parameterKey = ''Param_WorkingDaysPattern'';
					
					IF @sWorkingPattern IS null SET @fOK = 0;
				END
		
				IF @iAbsence_WorkingDaysType = 2	/* The Working Days are a numeric field reference. */
				BEGIN
					SELECT @sWorkingDaysNum_ColumnName = ASRSysColumns.columnName
					FROM ASRSysModuleSetup
					INNER JOIN ASRSysColumns
						ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
					WHERE moduleKey = ''MODULE_ABSENCE''
					AND parameterKey = ''Param_FieldWorkingDays'';
		
					IF @sWorkingDaysNum_ColumnName IS null SET @fOK = 0;
		
					IF @fOK = 1
					BEGIN
						SET @sParamDefinition = N''@workingDays varchar(MAX) OUTPUT''
						SET @sCommandString = ''SELECT @workingDays = '' + @sWorkingDaysNum_ColumnName + 
							'' FROM '' + @sPersonnelTableName + 
							'' WHERE id = '' + convert(varchar(128), @iPersonnelRecordID);
						EXECUTE sp_executesql @sCommandString, @sParamDefinition, @iWorkingDaysPerWeek OUTPUT;
		
						IF (@iWorkingDaysPerWeek IS null) SET @fOK = 0;
					END
				END
		
				IF @iAbsence_WorkingDaysType = 3	/* The Working Days are an working pattern field. */
				BEGIN
					SELECT @sWorkingDaysPattern_ColumnName = ASRSysColumns.columnName
					FROM ASRSysModuleSetup
					INNER JOIN ASRSysColumns
						ON convert(integer, ASRSysModuleSetup.parameterValue) = ASRSysColumns.columnID
					WHERE moduleKey = ''MODULE_ABSENCE''
					AND parameterKey = ''Param_FieldWorkingDays'';
		
					IF @sWorkingDaysPattern_ColumnName IS null SET @fOK = 0;
		
					IF @fOK = 1
					BEGIN
						SET @sParamDefinition = N''@workingDaysPattern varchar(MAX) OUTPUT''
						SET @sCommandString = ''SELECT @workingDaysPattern = '' + @sWorkingDaysNum_ColumnName + 
							'' FROM '' + @sPersonnelTableName + 
							'' WHERE id = '' + convert(varchar(128), @iPersonnelRecordID);
						EXECUTE sp_executesql @sCommandString, @sParamDefinition, @sWorkingPattern OUTPUT;
		
						IF (@sWorkingPattern IS null) SET @fOK = 0;
					END
				END
		
				IF @fOK = 1
				BEGIN
					/* Calculate the number of qualifying days per week. */
					IF len(@sWorkingPattern) > 0
					BEGIN
						SET @iLoop = 1;
		
						WHILE (len(@sWorkingPattern) >= (@iLoop * 2)) AND (@iLoop <=14)
						BEGIN
							IF (substring(@sWorkingPattern, @iLoop, 1) <> '' '') AND (substring(@sWorkingPattern, @iLoop + 1, 1) <> '' '')
							BEGIN
								SET @iWorkingDaysPerWeek = @iWorkingDaysPerWeek + 1;
							END
						
							SET @iLoop = @iLoop + 2;
						END
					END
		
					IF @iWorkingDaysPerWeek <= 0 SET @fOK = 0;
				END
			END
		
			IF @fOK = 1
			BEGIN
				SET @iConsecutiveRecords = 0;
				SET @dtLastWholeEndDate = null;
		
				/* Create a cursor of the absence records for the current person. */
				SET @sParamDefinition = N''@absenceRecs cursor OUTPUT''
				SET @sCommandString = ''SET @absenceRecs = CURSOR  LOCAL FAST_FORWARD FOR'' +
					'' SELECT '' + @sAbsenceTableName + ''.id, '' + 
						''convert(datetime, convert(varchar(20), '' + @sAbsenceTableName + ''.'' + @sAbsence_StartDateColumnName + '', 101)), '' + 
						''convert(datetime, convert(varchar(20), '' + @sAbsenceTableName + ''.'' + @sAbsence_EndDateColumnName + '', 101)), '' +
						''upper(left('' + @sAbsenceTableName + ''.'' + @sAbsence_StartSessionColumnName + '', 2)), '' +
						''upper(left('' + @sAbsenceTableName + ''.'' + @sAbsence_EndSessionColumnName + '', 2)), '' + 
						@sAbsenceTableName + ''.'' + @sAbsence_SSPAppliesColumnName + '', '' +
						@sAbsenceTableName + ''.'' + @sAbsence_QualifyingDaysColumnName + '', '' +
						@sAbsenceTableName + ''.'' + @sAbsence_WaitingDaysColumnName + '', '' +
						@sAbsenceTableName + ''.'' + @sAbsence_PaidDaysColumnName + 
					'' FROM '' + @sAbsenceTableName + 
					'' INNER JOIN '' + @sAbsenceTypeTableName + '' ON '' + @sAbsenceTableName + ''.'' + @sAbsence_TypeColumnName + '' = '' + @sAbsenceTypeTableName + ''.'' + @sAbsenceType_TypeColumnName +
					'' WHERE '' + @sAbsenceTableName + ''.id_'' + convert(varchar(128), @iPersonnelTableID) + '' = '' + convert(varchar(128), @iPersonnelRecordID) +
					'' AND '' + @sAbsenceTypeTableName + ''.'' + @sAbsenceType_SSPAppliesColumnName + '' = 1'' +
					'' ORDER BY '' + @sAbsenceTableName + ''.'' + @sAbsence_StartDateColumnName + '', '' + @sAbsenceTableName + ''.id'' +
					'' OPEN @absenceRecs'';
				EXECUTE sp_executesql @sCommandString, @sParamDefinition, @cursAbsenceRecords OUTPUT;
		
				/* Loop through the absence records, calculating SSP for each record. 
				NB. We check if any periods of absence are consecutive before checking for SSP application. */
				FETCH NEXT FROM @cursAbsenceRecords INTO @iAbsenceRecordID, @dtStartDate, @dtEndDate, @sStartSession, @sEndSession, @fOriginalSSPApplies, @dblOriginalQualifyingDays, @dblOriginalWaitingDays, @dblOriginalPaidDays;
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Ignore incomplete absence records. */
					IF (NOT @dtStartDate IS null) AND (NOT @dtEndDate IS null)
					BEGIN
						/* Ignore absence after retirement. */
						IF NOT @dtRetirementDate IS null
						BEGIN
							IF (@dtRetirementDate < @dtEndDate) 
							BEGIN
								SET @dtEndDate = @dtRetirementDate;
								SET @sEndSession = ''PM'';
							END
						END
						/* Ignore absence before the sixteenth birthday. */
						IF NOT @dtSixteenthBirthday IS null
						BEGIN
							IF (@dtSixteenthBirthday > @dtStartDate) 
							BEGIN
								SET @dtStartDate = @dtSixteenthBirthday;
								SET @sStartSession = ''AM'';
							END
						END
		
						/* Get the start and end dates (whole days only) of the current absence record. */
						SET @dtWholeStartDate = @dtStartDate;
						SET @dtWholeEndDate = @dtEndDate;
						IF @sStartSession = ''PM'' SET @dtWholeStartDate = @dtWholeStartDate + 1;
						IF @sEndSession = ''AM'' SET @dtWholeEndDate = @dtWholeEndDate - 1;
		
						IF @iConsecutiveRecords = 0 
						BEGIN
							SET @dtConsecutiveStartDate = @dtStartDate;
							SET @dtConsecutiveEndDate = @dtEndDate;
							SET @sConsecutiveStartSession = @sStartSession;
							SET @sConsecutiveEndSession = @sEndSession;
							SET @dtConsecutiveWholeStartDate = @dtWholeStartDate;
							SET @dtConsecutiveWholeEndDate = @dtWholeEndDate;
		
							/* Create a cursor of the absence records for the current person that follow the current absence record. */
							SET @sParamDefinition = N''@followingAbsenceRecs cursor OUTPUT'';
							SET @sCommandString = ''SET @followingAbsenceRecs = CURSOR  LOCAL FAST_FORWARD FOR'' +
								'' SELECT convert(datetime, convert(varchar(20), '' + @sAbsenceTableName + ''.'' + @sAbsence_StartDateColumnName + '', 101)), '' + 
									''convert(datetime, convert(varchar(20), '' + @sAbsenceTableName + ''.'' + @sAbsence_EndDateColumnName + '', 101)), '' +
									''upper(left('' + @sAbsenceTableName + ''.'' + @sAbsence_StartSessionColumnName + '', 2)), '' +
									''upper(left('' + @sAbsenceTableName + ''.'' + @sAbsence_EndSessionColumnName + '', 2)) '' + 
								'' FROM '' + @sAbsenceTableName + 
								'' INNER JOIN '' + @sAbsenceTypeTableName + '' ON '' + @sAbsenceTableName + ''.'' + @sAbsence_TypeColumnName + '' = '' + @sAbsenceTypeTableName + ''.'' + @sAbsenceType_TypeColumnName +
								'' WHERE '' + @sAbsenceTableName + ''.id_'' + convert(varchar(128), @iPersonnelTableID) + '' = '' + convert(varchar(128), @iPersonnelRecordID) +
								'' AND '' + @sAbsenceTypeTableName + ''.'' + @sAbsenceType_SSPAppliesColumnName + '' = 1'' +
								'' AND (NOT '' + @sAbsenceTableName + ''.'' + @sAbsence_StartDateColumnName + '' IS null)'' + 
								'' AND (NOT '' + @sAbsenceTableName + ''.'' + @sAbsence_EndDateColumnName + '' IS null)'' +
								'' AND ((convert(varchar(20), '' + @sAbsenceTableName + ''.'' + @sAbsence_StartDateColumnName + '', 112) > '' + convert(varchar(20), @dtStartDate, 112) + '')'' +
								'' OR ((convert(varchar(20), '' + @sAbsenceTableName + ''.'' + @sAbsence_StartDateColumnName + '', 112) = '' + convert(varchar(20), @dtStartDate, 112) + '') AND ('' + @sAbsenceTableName + ''.id > '' + convert(varchar(128), @iAbsenceRecordID) + '')))'' +
								'' ORDER BY '' + @sAbsenceTableName + ''.'' + @sAbsence_StartDateColumnName + '', '' + @sAbsenceTableName + ''.id'' +
								'' OPEN @followingAbsenceRecs'';
							EXECUTE sp_executesql @sCommandString, @sParamDefinition, @cursFollowingAbsenceRecords OUTPUT;
		
							SET @fContinue = 1;
							FETCH NEXT FROM @cursFollowingAbsenceRecords INTO @dtFollowingStartDate, @dtFollowingEndDate, @sFollowingStartSession, @sFollowingEndSession;
							WHILE (@@fetch_status = 0) AND (@fContinue = 1)
							BEGIN
								SET @fContinue = 0;
					
								/* Get the start and end dates (whole days only) of the current absence records. */
								SET @dtFollowingWholeStartDate = @dtFollowingStartDate;
								SET @dtFollowingWholeEndDate = @dtFollowingEndDate;
								IF @sFollowingStartSession = ''PM'' SET @dtFollowingWholeStartDate = @dtFollowingWholeStartDate + 1;
								IF @sFollowingEndSession = ''AM'' SET @dtFollowingWholeEndDate = @dtFollowingWholeEndDate - 1;
		
								IF ((@dtConsecutiveEndDate = @dtFollowingStartDate) AND (@sConsecutiveEndSession = ''AM'') AND (@sFollowingStartSession = ''PM''))
									OR (@dtConsecutiveWholeEndDate + 1 >= @dtFollowingWholeStartDate)
								BEGIN
									SET @iConsecutiveRecords = @iConsecutiveRecords + 1;
									SET @dtConsecutiveEndDate = @dtFollowingEndDate;
									SET @sConsecutiveEndSession = @sFollowingEndSession;
									SET @dtConsecutiveWholeEndDate = @dtFollowingWholeEndDate;
									SET @fContinue = 1;
								END
		
								FETCH NEXT FROM @cursFollowingAbsenceRecords INTO @dtFollowingStartDate, @dtFollowingEndDate, @sFollowingStartSession, @sFollowingEndSession;
							END
		
							CLOSE @cursFollowingAbsenceRecords;
							DEALLOCATE @cursFollowingAbsenceRecords;
		
						END
						ELSE
						BEGIN
							SET @iConsecutiveRecords = @iConsecutiveRecords - 1;
						END
		
						/* SSP Applies if the absence period is greater than 3 days. */
						SET @fSSPApplies = 0;
						IF (datediff(dd, @dtConsecutiveWholeStartDate, @dtConsecutiveWholeEndDate) + 1) > 3 SET @fSSPApplies = 1;
		
						IF @fSSPApplies = 1
						BEGIN
							/* Check if 56 days have passed since the previous absence period. */
							IF @dtLastWholeEndDate IS null
							BEGIN
								/* First absence record so use default values. */
								SET @dblWaitEntitlement = 3;
								SET @dblAbsenceEntitlement = @iWorkingDaysPerWeek * 28;
								SET @dtFirstLinkedWholeStartDate = @dtWholeStartDate;
							END
							ELSE
							BEGIN
								IF (datediff(dd, @dtLastWholeEndDate, @dtWholeStartDate) - 1) > 56
								BEGIN
									/* More than 56 days since the previous absence record so use default values. */
									SET @dblWaitEntitlement = 3;
									SET @dblAbsenceEntitlement = @iWorkingDaysPerWeek * 28;
									SET @dtFirstLinkedWholeStartDate = @dtWholeStartDate;
								END
							END
				
							/* Calculate SSP qualifying, waiting and paid days.
							NB. The start and end dates should already take into account the start and end periods (AM/PM)
							so that only whole absence days are used. */
							SET @dblQualifyingDays = 0;
		
							/* Loop from the start date to the end date, incrementing the number of qualifying days for each date that qualifies. */
							SET @dtTempDate = @dtStartDate;
		
							WHILE (@dtTempDate <= @dtEndDate)
							BEGIN
								SET @fAddOK = 0;
								SET @dblAddAmount = 0;
		
								IF len(@sWorkingPattern) = 0
								BEGIN
									/* No working pattern passed in, so use the ''daysPerWeek'' variable. */
									IF (@iWorkingDaysPerWeek = 7) OR 
										((datepart(dw, @dtTempDate) >= 2) AND (datepart(dw, @dtTempDate) <= 6))
									BEGIN
										/* The current date qualifies if 7 days per week are worked, or if the current date is a weekday. */
										SET @fAddOK = 1;
									END
								END
								ELSE	
								BEGIN
									/* Use the working pattern. */
									SET @iIndex = (2 * datepart(dw, @dtTempDate)) -1;
									IF len(@sWorkingPattern) >= (@iIndex +1)
									BEGIN
										/* The current date qualifies if its ''day of the week'' is worked in the working pattern.
										NB. Both AM and PM sessions must be worked for the day to qualify. */
										IF (substring(@sWorkingPattern, @iIndex, 1) <> '' '') AND (substring(@sWorkingPattern, @iIndex + 1, 1) <> '' '')
										BEGIN
											SET @fAddOK = 1;
										END
									END
								END
		
								IF @fAddOK = 1 
								BEGIN
									/* If the person is older than retirement age, then the day does not qualify. */
									IF NOT @dtRetirementDate IS null
									BEGIN
										IF @dtTempDate > @dtRetirementDate SET @fAddOK = 0;
									END
								END
		
								IF @fAddOK = 1 
								BEGIN
									/* If the person is less than sixteen then the day does not qualify. */
									IF (NOT @dtSixteenthBirthday IS null) 
									BEGIN
										IF @dtTempDate < @dtSixteenthBirthday SET @fAddOK = 0;
									END
								END
		
								IF @fAddOK = 1 
								BEGIN
									/* Days linked after 3 years from the start of the link do not count. */
									exec sp_ASRFn_WholeYearsBetweenTwoDates @iYearDifference OUTPUT, @dtFirstLinkedWholeStartDate, @dtTempDate;
									IF @iYearDifference >= 3  SET @fAddOK = 0;
								END
		
								/* Calculate how much to add to the Qualifying Days. */
								IF @fAddOK = 1 
								BEGIN
									SET @dblAddAmount = 0;
		
									IF @dtTempDate < @dtWholeStartDate
									BEGIN
										/* The current date is the half day before the whole dated period starts.
										A half day qualifies only if this period of absence consecutively follows another. */
										IF (@dtConsecutiveStartDate < @dtStartDate) OR 
											((@dtConsecutiveStartDate = @dtStartDate) AND (@sConsecutiveStartSession <> @sStartSession)) SET @dblAddAmount = 0.5;
									END
									ELSE
									BEGIN
										IF @dtTempDate > @dtWholeEndDate
										BEGIN
											/* The current date is the half day after the whole dated period end.
											A half day qualifies only if this period of absence is consecutively followed by another. */
											IF (@dtConsecutiveEndDate > @dtEndDate) OR 
												((@dtConsecutiveEndDate = @dtEndDate) AND (@sConsecutiveEndSession <> @sStartSession)) SET @dblAddAmount = 0.5;
										END
										ELSE
										BEGIN
											/* The current date lies within the whole dated period, so a whole day qualifies. */
											SET @dblAddAmount = 1;
										END
									END
								END
		
		
								/* Increment the number of qualifying days. */
								SET @dblQualifyingDays = @dblQualifyingDays + @dblAddAmount;
		
								SET @dtTempDate = @dtTempDate + 1;
							END
		
							/* Take off any waiting entitlement. */
							IF @dblWaitEntitlement > @dblQualifyingDays
							BEGIN
								SET @dblWaitingDays = @dblQualifyingDays;
								SET @dblWaitEntitlement = @dblWaitEntitlement - @dblQualifyingDays;
							END
							ELSE
							BEGIN
								SET @dblWaitingDays = @dblWaitEntitlement;
								SET @dblWaitEntitlement = 0;
							END
		
							/* Paid days is the difference providing there is enough entitlement. */
							SET @dblPaidDays = @dblQualifyingDays - @dblWaitingDays;
		
							IF @dblPaidDays > @dblAbsenceEntitlement
							BEGIN
								SET @dblPaidDays = @dblAbsenceEntitlement;
								SET @dblAbsenceEntitlement = 0;
							END
							ELSE
							BEGIN
								SET @dblAbsenceEntitlement = @dblAbsenceEntitlement - @dblPaidDays;
							END	
		
							SET @dtLastWholeEndDate = @dtWholeEndDate;
		
							/* Update the SSP fields in the current absence record if required. */
							IF (@fOriginalSSPApplies IS null) OR
								(@fOriginalSSPApplies = 0) OR
								(@dblOriginalQualifyingDays IS null) OR
								(@dblOriginalQualifyingDays <> @dblQualifyingDays) OR
								(@dblOriginalWaitingDays IS null) OR
								(@dblOriginalWaitingDays <> @dblWaitingDays) OR
								(@dblOriginalPaidDays IS null) OR
								(@dblOriginalPaidDays <> @dblPaidDays)
							BEGIN
								SET @sCommandString = ''UPDATE '' + @sAbsenceTableName +
									'' SET '' + @sAbsence_SSPAppliesColumnName + '' = 1, '' +
									@sAbsence_QualifyingDaysColumnName + '' = '' + convert(varchar(MAX), @dblQualifyingDays) + '', '' +
									@sAbsence_WaitingDaysColumnName + '' = '' + convert(varchar(MAX), @dblWaitingDays) + '', '' +
									@sAbsence_PaidDaysColumnName + '' = '' + convert(varchar(MAX), @dblPaidDays) + 
									'' WHERE id = '' + convert(varchar(128), @iAbsenceRecordID);
								exec sp_executesql @sCommandString;
							END
						END
						ELSE			
						BEGIN
							/* Update the SSP fields in the current absence record. */
							IF (@fOriginalSSPApplies IS null) OR
								(@fOriginalSSPApplies = 1) OR
								(@dblOriginalQualifyingDays IS null) OR
								(@dblOriginalQualifyingDays <> 0) OR
								(@dblOriginalWaitingDays IS null) OR
								(@dblOriginalWaitingDays <> 0) OR
								(@dblOriginalPaidDays IS null) OR
								(@dblOriginalPaidDays <> 0)
							BEGIN
								SET @sCommandString = ''UPDATE '' + @sAbsenceTableName +
									'' SET '' + @sAbsence_SSPAppliesColumnName + '' = 0, '' +
									@sAbsence_QualifyingDaysColumnName + '' = 0, '' +
									@sAbsence_WaitingDaysColumnName + '' = 0, '' +
									@sAbsence_PaidDaysColumnName + '' = 0'' + 
									'' WHERE id = '' + convert(varchar(128), @iAbsenceRecordID);
								exec sp_executesql @sCommandString;
							END
						END
					END
					ELSE
					BEGIN
						/* Update the SSP fields in the current absence record. */
						IF (@fOriginalSSPApplies IS null) OR
							(@fOriginalSSPApplies = 1) OR
							(@dblOriginalQualifyingDays IS null) OR
							(@dblOriginalQualifyingDays <> 0) OR
							(@dblOriginalWaitingDays IS null) OR
							(@dblOriginalWaitingDays <> 0) OR
							(@dblOriginalPaidDays IS null) OR
							(@dblOriginalPaidDays <> 0)
						BEGIN
							SET @sCommandString = ''UPDATE '' + @sAbsenceTableName +
								'' SET '' + @sAbsence_SSPAppliesColumnName + '' = 0, '' +
								@sAbsence_QualifyingDaysColumnName + '' = 0, '' +
								@sAbsence_WaitingDaysColumnName + '' = 0, '' +
								@sAbsence_PaidDaysColumnName + '' = 0'' + 
								'' WHERE id = '' + convert(varchar(128), @iAbsenceRecordID);
							exec sp_executesql @sCommandString;
						END
					END
		
					FETCH NEXT FROM @cursAbsenceRecords INTO @iAbsenceRecordID, @dtStartDate, @dtEndDate, @sStartSession, @sEndSession, @fOriginalSSPApplies, @dblOriginalQualifyingDays, @dblOriginalWaitingDays, @dblOriginalPaidDays;
				END
				CLOSE @cursAbsenceRecords;
				DEALLOCATE @cursAbsenceRecords;
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRFn_SystemTime
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_SystemTime]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_SystemTime];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_SystemTime]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_SystemTime]
		(
			@psResult	varchar(MAX) OUTPUT
		)
		AS
		BEGIN
			SET @psResult = convert(varchar(20), GETDATE(), 8);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRGetAuditTrail
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRGetAuditTrail]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRGetAuditTrail];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRGetAuditTrail]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRGetAuditTrail] (
			@piAuditType	int,
			@psOrder 		varchar(MAX))
		AS
		BEGIN
		
			SET NOCOUNT ON;
		
			DECLARE @sSQL			varchar(MAX),
					@sExecString	nvarchar(MAX);
		
			IF @piAuditType = 1
			BEGIN
		
				SET @sSQL = ''SELECT ASRSysAuditTrail.userName AS [User], 
					ASRSysAuditTrail.dateTimeStamp AS [Date / Time], 
					ASRSysAuditTrail.tableName AS [Table], 
					ASRSysAuditTrail.columnName AS [Column], 
					ASRSysAuditTrail.oldValue AS [Old Value], 
					ASRSysAuditTrail.newValue AS [New Value], 
					ASRSysAuditTrail.recordDesc AS [Record Description],
					ASRSysAuditTrail.id
					FROM ASRSysAuditTrail '';
		
				IF LEN(@psOrder) > 0
					SET @sExecString = @sSQL + @psOrder;
				ELSE
					SET @sExecString = @sSQL;
				
			END
			ELSE IF @piAuditType = 2
			BEGIN
		
				SET @sSQL =  ''SELECT userName AS [User], 
					dateTimeStamp AS [Date / Time],
					groupName AS [User Group],
					viewTableName AS [View / Table],
					columnName AS [Column], 
					action AS [Action],
					permission AS [Permission], 
					id
					FROM ASRSysAuditPermissions '';
		
				IF LEN(@psOrder) > 0
					SET @sExecString = @sSQL + @psOrder;
				ELSE
					SET @sExecString = @sSQL;
		
			END
			ELSE IF @piAuditType = 3
			BEGIN
				SET @sSQL = ''SELECT userName AS [User],
		    			dateTimeStamp AS [Date / Time],
					groupName AS [User Group], 
					userLogin AS [User Login],
					[Action], 
					id
					FROM ASRSysAuditGroup '';
		
				IF LEN(@psOrder) > 0
					SET @sExecString = @sSQL + @psOrder;
				ELSE
					SET @sExecString = @sSQL;
		
			END
			ELSE IF @piAuditType = 4
			BEGIN
				SET @sSQL = ''SELECT DateTimeStamp AS [Date / Time],
		    			UserGroup AS [User Group],
					UserName AS [User], 
					ComputerName AS [Computer Name],
					HRProModule AS [HR Pro Module],
					Action AS [Action], 
					id
					FROM ASRSysAuditAccess '';
		
				IF LEN(@psOrder) > 0
					SET @sExecString = @sSQL + @psOrder;
				ELSE
					SET @sExecString = @sSQL;
		
			END
		
			-- Retreive selected data
			IF LEN(@sExecString) > 0 EXECUTE sp_executeSQL @sExecString;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRGetControlDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRGetControlDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRGetControlDetails];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRGetControlDetails]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRGetControlDetails] 
		(@piScreenID int)
	AS
	BEGIN
		SELECT cont.*, 
			col.[columnName], col.[columnType], col.[datatype], col.[defaultValue],
			col.[size], col.[decimals], col.[lookupTableID], 
			col.[lookupColumnID], col.[lookupFilterColumnID], col.[lookupFilterOperator], col.[lookupFilterValueID], 
			col.[spinnerMinimum], col.[spinnerMaximum], col.[spinnerIncrement], 
			col.[mandatory], col.[uniquecheck], col.[uniquechecktype], col.[convertcase], 
			col.[mask], col.[blankIfZero], col.[multiline], col.[alignment] AS colAlignment, 
			col.[calcExprID], col.[gotFocusExprID], col.[lostFocusExprID], col.[dfltValueExprID], col.[calcTrigger], 
			CASE WHEN ISNULL(col.readOnly,0) = 1 THEN 1 ELSE CASE WHEN ISNULL(cont.readOnly,0) = 1 THEN 1 ELSE 0 END END AS ''readOnly'', 
			col.[statusBarMessage], col.[errorMessage], col.[linkTableID], col.[linkViewID],
			col.[linkOrderID], col.[Afdenabled], tab.[TableName],col.[Trimming], col.[Use1000Separator],
			col.[QAddressEnabled], col.[OLEType], col.[MaxOLESizeEnabled], col.[MaxOLESize], col.[AutoUpdateLookupValues]
		FROM [dbo].[ASRSysControls] cont
			LEFT OUTER JOIN [dbo].[ASRSysTables] tab ON cont.[tableID] = tab.[tableID]
			LEFT OUTER JOIN [dbo].[ASRSysColumns] col ON col.[tableID] = cont.[tableID] AND col.[columnID] = cont.[columnID]
		WHERE cont.[ScreenID] = @piScreenID
		ORDER BY cont.[PageNo], 
			cont.[ControlLevel] DESC, 
			cont.[tabIndex];
	END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRGetHistoryScreens
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRGetHistoryScreens]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRGetHistoryScreens];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRGetHistoryScreens]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRGetHistoryScreens]
			(@piParentScreenID	integer)
		AS
		BEGIN
			/* Return a recordset of the history screens that hang off the given parent screen. */
			SELECT ASRSysTables.tableName, 
				ASRSysTables.tableID,
				childScreens.screenID,
				childScreens.name,
				childScreens.pictureID
			FROM ASRSysScreens parentScreen
			INNER JOIN ASRSysHistoryScreens 
				ON parentScreen.screenID = ASRSysHistoryScreens.parentScreenID
			INNER JOIN ASRSysScreens childScreens 
				ON ASRSysHistoryScreens.historyScreenID = childScreens.screenID
			INNER JOIN ASRSysTables 
				ON childScreens.tableID = ASRSysTables.tableID
			WHERE parentScreen.screenID = @piParentScreenID
				AND childScreens.quickEntry = 0;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRGetOrders
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRGetOrders]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRGetOrders];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRGetOrders]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;


	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRGetOrders] (
			@piViewID	int,
			@piTableID	int)
		AS
		BEGIN
			SELECT DISTINCT ASRSysOrders.orderID, 
				ASRSysOrders.name, 
				ASRSysOrders.tableID 
			FROM ASRSysOrders 
			INNER JOIN ASRSysOrderItems ON ASRSysOrders.orderID = ASRSysOrderItems.orderID 
			INNER JOIN ASRSysViewColumns ON ASRSysOrderItems.columnID = ASRSysViewColumns.columnID 
			WHERE ASRSysOrders.tableID = @piTableID  
				AND ASRSysViewColumns.viewID = @piViewID  
				AND ASRSysViewColumns.inView = 1;
		END'
	EXECUTE sp_executeSQL @sSPCode;


	----------------------------------------------------------------------
	-- sp_ASRGetPickListRecords
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRGetPickListRecords]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRGetPickListRecords];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRGetPickListRecords]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;


	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRGetPickListRecords] (
			@piPickListID int)
		AS
		BEGIN
			SELECT ASRSysPickListItems.recordID AS id
			FROM ASRSysPickListItems 
			INNER JOIN ASRSysPickListName 
				ON ASRSysPickListItems.pickListID = ASRSysPickListName.pickListID
			WHERE ASRSysPickListName.pickListID = @piPickListID;
		END'
	EXECUTE sp_executeSQL @sSPCode;


	----------------------------------------------------------------------
	-- sp_ASRInsertNewUtility
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRInsertNewUtility]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRInsertNewUtility];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRInsertNewUtility]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRInsertNewUtility]
		(
		    @piNewRecordID	integer OUTPUT,   /* Output variable to hold the new record ID. */
		    @psInsertString nvarchar(MAX),    /* SQL Insert string to insert the new record. */
		    @psTableName	varchar(255),		 /* Table Name you want to retrieve */
		    @psIDColumnName varchar(30)      /* Name of the ID column  */
		)
		AS
		BEGIN
		    DECLARE @sCommand		nvarchar(MAX),
				@sParamDefinition 	nvarchar(MAX);
		
		    BEGIN TRANSACTION;
		
		    /* Run the given SQL INSERT string. */
		    EXECUTE sp_ExecuteSQL @psInsertString;
		
		    /* Get the ID of the inserted record.
		    NB. We do not use @@IDENTITY as the insertion that we have just performed may have triggered
		    other insertions (eg. into the Audit Trail table. The @@IDENTITY variable would then be the last IDENTITY value
		    entered in the Audit Trail table.*/
		    SET @sCommand = ''SELECT @recordID = MAX('' + @psIDColumnName + '') FROM '' + @psTableName + '';'';
		
		    SET @sParamDefinition = N''@recordID integer OUTPUT'';
		    EXEC sp_executesql @sCommand,  @sParamDefinition, @piNewRecordID OUTPUT;
		
		    COMMIT TRANSACTION;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_And
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_And]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_And];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_And]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_And]
		(	
			@pfResult	bit OUTPUT,
			@pfFirst	bit,
			@pfSecond	bit
		)
		AS
		BEGIN
			SET @pfResult = @pfFirst & @pfSecond;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_ConcatenatedWith
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_ConcatenatedWith]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_ConcatenatedWith];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_ConcatenatedWith]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_ConcatenatedWith] 
		(
			@psResult 	varchar(MAX) OUTPUT,
			@psString1 	varchar(MAX),
			@psString2	varchar(MAX)
		)
		AS
		BEGIN
			SET @psResult = @psString1 + @psString2;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsContainedWithin
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsContainedWithin]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsContainedWithin];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsContainedWithin]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsContainedWithin]
		(
			@pfResult   		bit OUTPUT,
			@psSearchString 	varchar(MAX),
			@psWholeString   	varchar(MAX)
		)
		AS
		BEGIN
			DECLARE @iTemp integer;
		
			SET @iTemp = charindex(@psSearchString, @psWholeString);
		
			IF @iTemp > 0
			BEGIN
				SET @pfResult = 1;
			END
			ELSE
			BEGIN
				SET @pfResult = 0;
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsEqualTo
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsEqualTo]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsEqualTo];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsEqualTo]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsEqualTo]
		(
			@date1			datetime,
			@date2			datetime,
			@retdate   		bit			OUTPUT,
			@char1			varchar(MAX),
			@char2			varchar(MAX),
			@retchar   		bit			OUTPUT,
			@numeric1		numeric,
			@numeric2		numeric,
			@retnumeric   	bit			OUTPUT,
			@logic1			bit,
			@logic2			bit,
			@retlogic		bit			OUTPUT
		)
		AS
		BEGIN
			if @date1 is not null
			begin
				if @date1 = @date2
				begin
				set @retdate = 1
				select @retdate as result
				end
				if @date1 <> @date2
				begin
				set @retdate = 0
				select @retdate as result
				end	
			end
		
			if @char1 is not null
			begin
				if @char1 = @char2
				begin
				set @retchar = 1
				select @retchar as result
				end
				if @char1 <> @char2
				begin
				set @retchar = 0
				select @retchar as result
				end	
			end
		
			if @numeric1 is not null
			begin
				if @numeric1 = @numeric2
				begin
				set @retnumeric = 1
				select @retnumeric as result
				end
				if @numeric1 <> @numeric2
				begin
				set @retnumeric = 0
				select @retnumeric as result
				end	
			end
		
			if @logic1 is not null
			begin
				if @logic1 = @logic2
				begin
				set @retlogic = 1
				select @retlogic as result
				end
				if @logic1 <> @logic2
				begin
				set @retlogic = 0
				select @retlogic as result
				end	
			end
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsEqualTo_1_1
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsEqualTo_1_1]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsEqualTo_1_1];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsEqualTo_1_1]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsEqualTo_1_1]
		(
			@pfResult   	bit OUTPUT,
			@psString1		varchar(MAX),
			@psString2		varchar(MAX)
		)
		AS
		BEGIN
			IF @psString1 = @psString2
			BEGIN
				SET @pfResult = 1;
			END
			ELSE
			BEGIN
				SET @pfResult = 0;
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsGreaterThan
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsGreaterThan]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsGreaterThan];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsGreaterThan]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsGreaterThan]
		(
			@date1			datetime,
			@date2			datetime,
			@retdate   		bit OUTPUT,
			@char1			varchar(MAX),
			@char2			varchar(MAX),
			@retchar   		bit OUTPUT,
			@numeric1		numeric,
			@numeric2		numeric,
			@retnumeric   	bit OUTPUT,
			@logic1			bit,
			@logic2			bit,
			@retlogic		bit OUTPUT
		)
		
		AS
		BEGIN
			if @date1 is not null
			begin
				if @date1 > @date2
				begin
				set @retdate = 1
				select @retdate as result
				end
				if @date1 <= @date2
				begin
				set @retdate = 0
				select @retdate as result
				end	
			end
		
			if @char1 is not null
			begin
				if @char1 > @char2
				begin
				set @retchar = 1
				select @retchar as result
				end
				if @char1 <= @char2
				begin
				set @retchar = 0
				select @retchar as result
				end	
			end
		
			if @numeric1 is not null
			begin
				if @numeric1 > @numeric2
				begin
				set @retnumeric = 1
				select @retnumeric as result
				end
				if @numeric1 <= @numeric2
				begin
				set @retnumeric = 0
				select @retnumeric as result
				end	
			end
		
			if @logic1 is not null
			begin
				if @logic1 > @logic2
				begin
				set @retlogic = 1
				select @retlogic as result
				end
				if @logic1 <= @logic2
				begin
				set @retlogic = 0
				select @retlogic as result
				end	
			end
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsGreaterThan_1_1
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsGreaterThan_1_1]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsGreaterThan_1_1];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsGreaterThan_1_1]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsGreaterThan_1_1]
		(
			@pfResult  	bit OUTPUT,
			@psString1	varchar(MAX),
			@psString2	varchar(MAX)
		)
		AS
		BEGIN
			IF @psString1 > @psString2
			BEGIN
				SET @pfResult = 1;
			END
			ELSE
			BEGIN
				SET @pfResult = 0;
			END	
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsGreaterThanOrEqualTo
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsGreaterThanOrEqualTo]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsGreaterThanOrEqualTo];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsGreaterThanOrEqualTo]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsGreaterThanOrEqualTo]
		(
			@date1			datetime,
			@date2			datetime,
			@retdate   		bit				OUTPUT,
			@char1			varchar(MAX),
			@char2			varchar(MAX),
			@retchar   		bit				OUTPUT,
			@numeric1		numeric,
			@numeric2		numeric,
			@retnumeric   	bit				OUTPUT,
			@logic1			bit,
			@logic2			bit,
			@retlogic		bit				OUTPUT
		)
		AS
		BEGIN
			if @date1 is not null
			begin
				if @date1 >= @date2
				begin
				set @retdate = 1
				select @retdate as result
				end
				if @date1 < @date2
				begin
				set @retdate = 0
				select @retdate as result
				end	
			end
		
			if @char1 is not null
			begin
				if @char1 >= @char2
				begin
				set @retchar = 1
				select @retchar as result
				end
				if @char1 < @char2
				begin
				set @retchar = 0
				select @retchar as result
				end	
			end
		
			if @numeric1 is not null
			begin
				if @numeric1 >= @numeric2
				begin
				set @retnumeric = 1
				select @retnumeric as result
				end
				if @numeric1 < @numeric2
				begin
				set @retnumeric = 0
				select @retnumeric as result
				end	
			end
		
			if @logic1 is not null
			begin
				if @logic1 >= @logic2
				begin
				set @retlogic = 1
				select @retlogic as result
				end
				if @logic1 < @logic2
				begin
				set @retlogic = 0
				select @retlogic as result
				end	
			end
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsGreaterThanOrEqualTo_1_1
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsGreaterThanOrEqualTo_1_1]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsGreaterThanOrEqualTo_1_1];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsGreaterThanOrEqualTo_1_1]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsGreaterThanOrEqualTo_1_1]
		(
			@pfResult  	bit OUTPUT,
			@psString1	varchar(MAX),
			@psString2	varchar(MAX)
		)
		AS
		BEGIN
			IF @psString1 >= @psString2
			BEGIN
				SET @pfResult = 1;
			END
			ELSE
			BEGIN
				SET @pfResult = 0;
			END	
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsLessThan
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsLessThan]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsLessThan];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsLessThan]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsLessThan]
		(
			@date1			datetime,
			@date2			datetime,
			@retdate   		bit				OUTPUT,
			@char1			varchar(MAX),
			@char2			varchar(MAX),
			@retchar   		bit OUTPUT,
			@numeric1		numeric,
			@numeric2		numeric,
			@retnumeric   	bit				OUTPUT,
			@logic1			bit,
			@logic2			bit,
			@retlogic		bit				OUTPUT
		)
		AS
		BEGIN
		
			if @date1 is not null
			begin
				if @date1 < @date2
				begin
				set @retdate = 1
				select @retdate as result
				end
				if @date1 >= @date2
				begin
				set @retdate = 0
				select @retdate as result
				end	
			end
		
			if @char1 is not null
			begin
				if @char1 < @char2
				begin
				set @retchar = 1
				select @retchar as result
				end
				if @char1 >= @char2
				begin
				set @retchar = 0
				select @retchar as result
				end	
			end
		
			if @numeric1 is not null
			begin
				if @numeric1 < @numeric2
				begin
				set @retnumeric = 1
				select @retnumeric as result
				end
				if @numeric1 >= @numeric2
				begin
				set @retnumeric = 0
				select @retnumeric as result
				end	
			end
		
			if @logic1 is not null
			begin
				if @logic1 < @logic2
				begin
				set @retlogic = 1
				select @retlogic as result
				end
				if @logic1 >= @logic2
				begin
				set @retlogic = 0
				select @retlogic as result
				end	
			end
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsLessThan_1_1
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsLessThan_1_1]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsLessThan_1_1];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsLessThan_1_1]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsLessThan_1_1]
		(
			@pfResult  	bit OUTPUT,
			@psString1	varchar(MAX),
			@psString2	varchar(MAX)
		)
		AS
		BEGIN
			IF @psString1 < @psString2
			BEGIN
				SET @pfResult = 1;
			END
			ELSE
			BEGIN
				SET @pfResult = 0;
			END	
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsLessThanOrEqualTo
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsLessThanOrEqualTo]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsLessThanOrEqualTo];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsLessThanOrEqualTo]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsLessThanOrEqualTo]
		(
			@date1			datetime,
			@date2			datetime,
			@retdate   		bit			OUTPUT,
			@char1			varchar(MAX),
			@char2			varchar(MAX),
			@retchar   		bit			OUTPUT,
			@numeric1		numeric,
			@numeric2		numeric,
			@retnumeric   	bit			OUTPUT,
			@logic1			bit,
			@logic2			bit,
			@retlogic		bit			OUTPUT
		)
		AS
		BEGIN
			if @date1 is not null
			begin
				if @date1 <= @date2
				begin
				set @retdate = 1
				select @retdate as result
				end
				if @date1 > @date2
				begin
				set @retdate = 0
				select @retdate as result
				end	
			end
		
			if @char1 is not null
			begin
				if @char1 <= @char2
				begin
				set @retchar = 1
				select @retchar as result
				end
				if @char1 > @char2
				begin
				set @retchar = 0
				select @retchar as result
				end	
			end
		
			if @numeric1 is not null
			begin
				if @numeric1 <= @numeric2
				begin
				set @retnumeric = 1
				select @retnumeric as result
				end
				if @numeric1 > @numeric2
				begin
				set @retnumeric = 0
				select @retnumeric as result
				end	
			end
		
			if @logic1 is not null
			begin
				if @logic1 <= @logic2
				begin
				set @retlogic = 1
				select @retlogic as result
				end
				if @logic1 > @logic2
				begin
				set @retlogic = 0
				select @retlogic as result
				end	
			end
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsLessThanOrEqualTo_1_1
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsLessThanOrEqualTo_1_1]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsLessThanOrEqualTo_1_1];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsLessThanOrEqualTo_1_1]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsLessThanOrEqualTo_1_1]
		(
			@pfResult	bit OUTPUT,
			@psString1	varchar(MAX),
			@psString2	varchar(MAX)
		)
		AS
		BEGIN
			IF @psString1 <= @psString2
			BEGIN
				SET @pfResult = 1;
			END
			ELSE
			BEGIN
				SET @pfResult = 0;
			END	
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsNotEqualTo
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsNotEqualTo]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsNotEqualTo];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsNotEqualTo]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsNotEqualTo]
		(
			@date1			datetime,
			@date2			datetime,
			@retdate   		bit			OUTPUT,
			@char1			varchar(MAX),
			@char2			varchar(MAX),
			@retchar   		bit			OUTPUT,
			@numeric1		numeric,
			@numeric2		numeric,
			@retnumeric   	bit			OUTPUT,
			@logic1			bit,
			@logic2			bit,
			@retlogic		bit			OUTPUT
		)
		AS
		BEGIN
			if @date1 is not null
			begin
				if @date1 = @date2
				begin
				set @retdate = 0
				select @retdate as result
				end
				if @date1 <> @date2
				begin
				set @retdate = 1
				select @retdate as result
				end	
			end
		
			if @char1 is not null
			begin
				if @char1 = @char2
				begin
				set @retchar = 0
				select @retchar as result
				end
				if @char1 <> @char2
				begin
				set @retchar = 1
				select @retchar as result
				end	
			end
		
			if @numeric1 is not null
			begin
				if @numeric1 = @numeric2
				begin
				set @retnumeric = 0
				select @retnumeric as result
				end
				if @numeric1 <> @numeric2
				begin
				set @retnumeric = 1
				select @retnumeric as result
				end	
			end
		
			if @logic1 is not null
			begin
				if @logic1 = @logic2
				begin
				set @retlogic = 0
				select @retlogic as result
				end
				if @logic1 <> @logic2
				begin
				set @retlogic = 1
				select @retlogic as result
				end	
			end
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_IsNotEqualTo_1_1
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_IsNotEqualTo_1_1]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_IsNotEqualTo_1_1];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_IsNotEqualTo_1_1]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_IsNotEqualTo_1_1]
		(
			@pfResult	bit OUTPUT,
			@psString1	varchar(MAX),
			@psString2	varchar(MAX)
		)
		AS
		BEGIN
			IF @psString1 = @psString2
			BEGIN
				SET @pfResult = 0;
			END
			ELSE
			BEGIN
				SET @pfResult = 1;
			END	
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASROp_Or
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASROp_Or]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASROp_Or];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASROp_Or]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASROp_Or]
		(
			@pfResult	bit OUTPUT,
			@pfFirst	bit,
			@pfSecond	bit
		)
		AS
		BEGIN
			SET @pfResult = @pfFirst | @pfSecond;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRPurgeDate
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRPurgeDate]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRPurgeDate];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRPurgeDate]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRPurgeDate]
		(
		    @purgedate varchar(MAX) OUTPUT,
		    @purgekey varchar(MAX)
		)
		AS
		BEGIN
		    DECLARE @unit char(1),
		            @period int,
		            @lastPurge datetime,
		            @today datetime;
		
		    /* Only get date and not current time */
		    select @today = convert(datetime,convert(varchar,getdate(),101));
		
		    /* Get purge period details */
		    SELECT @unit = unit
		         , @period = (period * -1)
		         , @lastPurge = lastpurgedate
		    FROM   asrsyspurgeperiods
		    WHERE  purgekey = @purgekey;
		
		    /* calculate purge date */
		    SELECT @purgedate = CASE @unit
		        WHEN ''D'' THEN dateadd(dd,@period,@today)
		        WHEN ''W'' THEN dateadd(ww,@period,@today)
		        WHEN ''M'' THEN dateadd(mm,@period,@today)
		        WHEN ''Y'' THEN dateadd(yy,@period,@today)
		    END;
		
		    IF @purgedate IS NULL OR datediff(d,@purgedate,@lastPurge) > 0
		    BEGIN
		      SET @purgedate = @lastPurge;
		    END
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRPurgeRecords
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRPurgeRecords]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRPurgeRecords];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRPurgeRecords]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRPurgeRecords]
		(
		    @PurgeKey varchar(255),
		    @TableName varchar(255),
		    @DateColumn varchar(255)
		)
		AS
		BEGIN
		
		    /* EXEC sp_ASRPurgeRecords ''EMAIL'', ''ASRSysEmailQueue'', ''DateDue'' */
		
		    DECLARE @PurgeDate datetime;
		    DECLARE @sSQL nvarchar(MAX);
		
		    EXEC [dbo].[sp_ASRPurgeDate] @PurgeDate OUTPUT, @PurgeKey;
		
		    SELECT @sSQL = ''DELETE FROM '' + @TableName + '' WHERE '' + @DateColumn + '' < '''''' + convert(varchar,@PurgeDate,101) + '''''''';
		    EXEC sp_executesql @sSQL;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRSendMessage
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRSendMessage]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRSendMessage];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRSendMessage]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRSendMessage] 
		(
			@psMessage	varchar(MAX),
			@psSPIDS	varchar(MAX)
		)
		AS
		BEGIN
			DECLARE @iDBid		integer,
				@iSPid			integer,
				@iUid			integer,
				@sLoginName		varchar(256),
				@dtLoginTime	datetime, 
				@sCurrentUser	varchar(256),
				@sCurrentApp	varchar(256),
				@Realspid		integer;
		
			CREATE TABLE #tblCurrentUsers				
				(
					hostname varchar(256)
					,loginame varchar(256)
					,program_name varchar(256)
					,hostprocess varchar(20)
					,sid binary(86)
					,login_time datetime
					,spid int
					,uid smallint);
					
			INSERT INTO #tblCurrentUsers
				EXEC spASRGetCurrentUsers;
		
			--MH20040224 Fault 8062
			--{
			--Need to get spid of parent process
			SELECT @Realspid = a.spid
			FROM #tblCurrentUsers a
			FULL OUTER JOIN #tblCurrentUsers b
				ON a.hostname = b.hostname
				AND a.hostprocess = b.hostprocess
				AND a.spid <> b.spid
			WHERE b.spid = @@Spid;
		
			--If there is no parent spid then use current spid
			IF @Realspid is null SET @Realspid = @@spid;
		
			/* Get the process information for the current user. */
			SELECT @iDBid = db_id(), 
				@sCurrentUser = loginame,
				@sCurrentApp = program_name
			FROM #tblCurrentUsers
			WHERE spid = @@Spid;
		
			/* Get a cursor of the other logged in HR Pro users. */
			DECLARE logins_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT spid, loginame, uid, login_time
				FROM #tblCurrentUsers
				WHERE (spid <> @@spid and spid <> @Realspid)
				AND (@psSPIDS = '''' OR charindex('' ''+convert(varchar,spid)+'' '', @psSPIDS)>0);
		
			OPEN logins_cursor;
			FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime;
			WHILE (@@fetch_status = 0)
			BEGIN
				/* Create a message record for each HR Pro user. */
				INSERT INTO ASRSysMessages 
					(loginname, [message], loginTime, [dbid], [uid], spid, messageTime, messageFrom, messageSource) 
					VALUES(@sLoginName, @psMessage, @dtLoginTime, @iDBid, @iUid, @iSPid, getdate(), @sCurrentUser, @sCurrentApp);
		
				FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime;
			END
			CLOSE logins_cursor;
			DEALLOCATE logins_cursor;
		
			IF OBJECT_ID(''tempdb..#tblCurrentUsers'', N''U'') IS NOT NULL
				DROP TABLE #tblCurrentUsers;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRSystemPermission
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRSystemPermission]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRSystemPermission];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRSystemPermission]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRSystemPermission]
		(
			@pfPermissionGranted 	bit OUTPUT,
			@psCategoryKey			varchar(50),
			@psPermissionKey		varchar(50),
			@psSQLLogin 			varchar(200)
		)
		AS
		BEGIN
			
			-- Return 1 if the given permission is granted to the current user, 0 if it is not.
			DECLARE @fGranted bit,
					@sGroupName varchar(255);
		
			-- Is logged in user a system administrator
			SELECT @fGranted = sysAdmin FROM master..syslogins WHERE loginname = @psSQLLogin;
		
			IF @fGranted = 0
			BEGIN
				SELECT @sGroupName = usg.name
				FROM sysusers usu
				left outer join
				(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
				WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) and
					(usg.issqlrole = 1 or usg.uid is null) and
					usu.name = @psSQLLogin AND not (usg.name like ''ASRSys%'')
					AND not (usg.name = ''db_owner'');
		
				SELECT @fGranted = ASRSysGroupPermissions.permitted
				FROM ASRSysGroupPermissions
					INNER JOIN ASRSysPermissionItems 
						ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID
					INNER JOIN ASRSysPermissionCategories
						ON ASRSysPermissionCategories.categoryID = ASRSysPermissionItems.categoryID
				WHERE ASRSysPermissionItems.itemKey = @psPermissionKey
					AND ASRSysGroupPermissions.groupName = @sGroupName
					AND ASRSysPermissionCategories.categoryKey = @psCategoryKey;
			END
		
		
			IF @fGranted IS NULL
			BEGIN
				SET @fGranted = 0;
			END
		
			SET @pfPermissionGranted = @fGranted;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- sp_ASRUniqueObjectName
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRUniqueObjectName]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRUniqueObjectName];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRUniqueObjectName]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRUniqueObjectName](
				  @psUniqueObjectName sysname OUTPUT
				, @Prefix sysname
				, @Type int)
		AS
		BEGIN
			DECLARE @NewObj 		as sysname
				, @Count 			as integer
				, @sUserName		as sysname
				, @sCommandString	nvarchar(MAX)	
		 		, @sParamDefinition	nvarchar(500);
		
			SET @sUserName = SYSTEM_USER;
			SET @Count = 1;
			SET @NewObj = @Prefix + CONVERT(varchar(100),@Count);
		
			WHILE (EXISTS (SELECT * FROM sysobjects WHERE id = object_id(@NewObj) AND sysstat & 0xf = @Type))
				OR (EXISTS (SELECT * FROM ASRSysSQLObjects WHERE Name = @NewObj AND Type = @Type))
				BEGIN
					SET @Count = @Count + 1;
					SET @NewObj = @Prefix + CONVERT(varchar(10),@Count);
				END
		
			INSERT INTO [dbo].[ASRSysSQLObjects] ([Name], [Type], [DateCreated], [Owner])
				VALUES (@NewObj, @Type, GETDATE(), @sUserName);
		
			SET @sCommandString = ''SELECT @psUniqueObjectName = '''''' + @NewObj + '''''''';
			SET @sParamDefinition = N''@psUniqueObjectName sysname output'';
			EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psUniqueObjectName output;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;


	----------------------------------------------------------------------
	-- spASRAccordNeedToSendAll
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRAccordNeedToSendAll]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRAccordNeedToSendAll];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRAccordNeedToSendAll]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRAccordNeedToSendAll] 
			(@iTransferType int, 
			@iRecordID int,
			@bResend bit OUTPUT)
		AS
		BEGIN
			SET NOCOUNT ON;
		
			DECLARE @Status integer;
		
			SELECT TOP 1 @Status = [Status] FROM [dbo].[ASRSysAccordTransactions]
				WHERE [HRProRecordID] = @iRecordID AND [TransferType] = @iTransferType
				ORDER BY [CreatedDateTime] DESC;
		
			-- Nothing found
			IF @Status IS NULL SET @bResend = 1;
		
			-- Previous transaction failed
			IF @Status IN (20) SET @bResend = 0;
		
			--	Previous transaction went as update - should be new
			IF @Status IN (22, 23, 31) SET @bResend = 1;
		
			-- Pending, success, or success with warnings, blocked
			IF @Status IN (1, 10, 11, 21, 30) SET @bResend = 0;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;


	----------------------------------------------------------------------
	-- spASRAccordPopulateTransaction
	----------------------------------------------------------------------
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRAccordPopulateTransaction]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRAccordPopulateTransaction];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRAccordPopulateTransaction]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRAccordPopulateTransaction] (
		@piTransactionID	integer OUTPUT,
		@piTransferType		integer,
		@piTransactionType	integer ,
		@piDefaultStatus	integer,
		@piHRProRecordID	integer,
		@iTriggerLevel		integer,
		@pbSendAllFields	bit OUTPUT)
	AS
	BEGIN	

		-- Return the required user or system setting.
		DECLARE @iCount			integer,
			@bNewTransaction	bit,
			@iStatus			integer,
			@bCreate			bit,
			@bForceAsUpdate		bit;

		SET @piTransactionID = null;
		SET @bCreate = 1;
		SET @bForceAsUpdate = 0;

		SELECT @piTransactionID = [TransactionID]
			FROM [dbo].[ASRSysAccordTransactionProcessInfo]
			WHERE [spid] = @@SPID AND [TransferType] = @piTransferType
				AND [RecordID] = @piHRProRecordID;

		-- Could be a null if the trigger was fired from a non Accord module enabled table, e.g. a child updating a parent field
		IF @piTransactionID IS null SET @bNewTransaction = 1;
		ELSE SET @bNewTransaction = 0;

		-- Get a transaction ID for this process and update the temporary Accord table
		IF @bNewTransaction = 1
		BEGIN
			SELECT @iCount = COUNT(*)
				FROM [dbo].[ASRSysSystemSettings]
				WHERE [section] = ''AccordTransfer'' AND [settingKey] = ''NextTransactionID'';
			
			IF @iCount = 0
				INSERT [dbo].[ASRSysSystemSettings] (Section, SettingKey, SettingValue)
					VALUES (''AccordTransfer'',''NextTransactionID'',1);
			ELSE
				UPDATE [dbo].[ASRSysSystemSettings] SET [SettingValue] = [SettingValue] + 1
					WHERE [section] = ''AccordTransfer'' AND [settingKey] =  ''NextTransactionID'';

			SELECT @piTransactionID = [settingValue]
				FROM[dbo].[ASRSysSystemSettings]
				WHERE [section] = ''AccordTransfer'' AND [settingKey] =  ''NextTransactionID'';

			-- If update, has it already been sent?
			IF @piTransactionType = 1
			BEGIN

				SELECT TOP 1 @iStatus = [Status] FROM [dbo].[ASRSysAccordTransactions]
				WHERE [HRProRecordID] = @piHRProRecordID
					AND [TransferType] = @piTransferType
				ORDER BY [CreatedDateTime] DESC;

				IF @iStatus IS NULL OR @iStatus = 23
				BEGIN
					SET @piTransactionType = 0;
					SET @pbSendAllFields = 1;
				END
				ELSE IF @iStatus = 20
				BEGIN
					IF EXISTS(SELECT [Status]
						FROM [dbo].[ASRSysAccordTransactions]
						WHERE [HRProRecordID] = @piHRProRecordID
							AND [Status] IN (10, 11) AND [TransferType] = @piTransferType)
					BEGIN
						SET @piTransactionType = 1;
					END
					ELSE
					BEGIN
						SET @piTransactionType = 0;
					END
					
					SET @pbSendAllFields = 1;
					
				END

			END

			SELECT @bForceAsUpdate = [ForceAsUpdate] FROM [dbo].[ASRSysAccordTransferTypes]
				WHERE [TransferTypeID] = @piTransferType;

			IF @bForceAsUpdate = 1 AND @piTransactionType = 0 SET @piTransactionType = 1;

			-- Are we trying to delete something thats never been sent?
			IF @piTransactionType = 2
			BEGIN
				SELECT TOP 1 @iStatus = [Status] FROM [dbo].[ASRSysAccordTransactions]
				WHERE [HRProRecordID] = @piHRProRecordID
				ORDER BY [CreatedDateTime] DESC;
			
				IF @iStatus IS NULL	SET @bCreate = 0;
				ELSE SET @pbSendAllFields = 1;
			END

			-- Insert a record into the Accord Transfer table.
			IF @bCreate = 1
			BEGIN
				INSERT INTO [dbo].[ASRSysAccordTransactions] ([TransactionID], [TransferType], [TransactionType], [CreatedUser], [CreatedDateTime], [Status], [HRProRecordID], [Archived])
					VALUES (@piTransactionID, @piTransferType, @piTransactionType, SYSTEM_USER, GETDATE(), @piDefaultStatus, @piHRProRecordID, 0);

				INSERT [dbo].[ASRSysAccordTransactionProcessInfo] ([SPID], [TransactionID], [TransferType], [RecordID])
					VALUES (@@SPID, @piTransactionID, @piTransferType, @piHRProRecordID);
			END

		END
	END';
	EXECUTE sp_executeSQL @sSPCode;




	----------------------------------------------------------------------
	-- spASRActionActiveWorkflowSteps
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRActionActiveWorkflowSteps]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRActionActiveWorkflowSteps];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
		AS
		BEGIN
			-- Return a recordset of the workflow steps that need to be actioned by the Workflow service.
			-- Action any that can be actioned immediately. 
			DECLARE
				@iAction			integer, -- 0 = do nothing, 1 = submit step, 2 = change status to ''2'', 3 = Summing Junction check, 4 = Or check
				@iElementType		integer,
				@iInstanceID		integer,
				@iElementID			integer,
				@iStepID			integer,
				@iCount				integer,
				@sStatus			bit,
				@sMessage			varchar(MAX),
				@iTemp				integer, 
				@iTemp2				integer, 
				@iTemp3				integer,
				@sForms 			varchar(MAX), 
				@iType				integer,
				@iDecisionFlow		integer,
				@fInvalidElements	bit, 
				@fValidElements		bit, 
				@iPrecedingElementID	integer, 
				@iPrecedingElementType	integer, 
				@iPrecedingElementStatus	integer, 
				@iPrecedingElementFlow	integer, 
				@fSaveForLater			bit;
		
			DECLARE stepsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT E.type,
				S.instanceID,
				E.ID,
				S.ID
			FROM ASRSysWorkflowInstanceSteps S
			INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
			WHERE S.status = 1
				AND E.type <> 5; -- 5 = StoredData elements handled in the service
		
			OPEN stepsCursor;
			FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID;
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
					END;
				
				IF @iAction = 3 -- Summing Junction check
				BEGIN
					-- Check if all preceding steps have completed before submitting this step.
					SET @fInvalidElements = 0;	
				
					DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT WE.ID,
						WE.type,
						WIS.status,
						WIS.decisionFlow
					FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
					INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
					INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
						AND WIS.instanceID = @iInstanceID;
		
					OPEN precedingElementsCursor;			
					FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
		
					WHILE (@@fetch_status = 0)
						AND (@fInvalidElements = 0)
					BEGIN
						IF (@iPrecedingElementType = 4) -- Decision
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
							BEGIN
								SELECT @iCount = COUNT(*) 
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
								WHERE ID = @iElementID;
		
								IF @iCount = 0 SET @fInvalidElements = 1;
							END
							ELSE
							BEGIN
								SET @fInvalidElements = 1;
							END
						END
						ELSE
						BEGIN
							IF (@iPrecedingElementType = 2) -- WebForm
							BEGIN
								IF @iPrecedingElementStatus = 3 -- 3 = completed
									OR @iPrecedingElementStatus = 6 -- 6 = timeout
								BEGIN
									SET @iTemp3 = CASE
											WHEN @iPrecedingElementStatus = 3 THEN 0
											ELSE 1
										END;
		
									SELECT @iCount = COUNT(*)
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
									WHERE ID = @iElementID;
								
									IF @iCount = 0 SET @fInvalidElements = 1;
								END
								ELSE
								BEGIN
									SET @fInvalidElements = 1;
								END
							END
							ELSE
							BEGIN
								IF (@iPrecedingElementType = 5) -- StoredData
								BEGIN
									IF @iPrecedingElementStatus = 3 -- 3 = completed
										OR @iPrecedingElementStatus = 8 -- 8 = failed action
									BEGIN
										SET @iTemp3 = CASE
												WHEN @iPrecedingElementStatus = 3 THEN 0
												ELSE 1
											END;
		
										SELECT @iCount = COUNT(*)
										FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
										WHERE ID = @iElementID;
									
										IF @iCount = 0 SET @fInvalidElements = 1;
									END
									ELSE
									BEGIN
										SET @fInvalidElements = 1;
									END
								END
								ELSE
								BEGIN
									-- Preceding element must have status 3 (3 =Completed)
									IF @iPrecedingElementStatus <> 3 SET @fInvalidElements = 1;
								END
							END
						END
		
						FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
					END
					CLOSE precedingElementsCursor;
					DEALLOCATE precedingElementsCursor;
					
					IF (@fInvalidElements = 0) SET @iAction = 1;
				END
		
				IF @iAction = 4 -- Or check
				BEGIN
					SET @fValidElements = 0;
					-- Check if any preceding steps have completed before submitting this step. 
		
					DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT WE.ID,
						WE.type,
						WIS.status,
						WIS.decisionFlow
					FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
					INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
					INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
						AND WIS.instanceID = @iInstanceID;
		
					OPEN precedingElementsCursor;	
		
					FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
					WHILE (@@fetch_status = 0)
						AND (@fValidElements = 0)
					BEGIN
						IF (@iPrecedingElementType = 4) -- Decision
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
							BEGIN
								SELECT @iCount = COUNT(*)
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
								WHERE ID = @iElementID;
							
								IF @iCount > 0 SET @fValidElements = 1;
							END
						END
						ELSE
						BEGIN
							IF (@iPrecedingElementType = 2) -- WebForm
							BEGIN
								IF @iPrecedingElementStatus = 3 -- 3 = completed
									OR @iPrecedingElementStatus = 6 -- 6 = timeout
								BEGIN
									SET @iTemp3 = CASE
											WHEN @iPrecedingElementStatus = 3 THEN 0
											ELSE 1
										END;
		
									SELECT @iCount = COUNT(*)
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
									WHERE ID = @iElementID;
							
									IF @iCount > 0 SET @fValidElements = 1;
								END
							END
							ELSE
							BEGIN
								IF (@iPrecedingElementType = 5) -- StoredData
								BEGIN
									IF @iPrecedingElementStatus = 3 -- 3 = completed
										OR @iPrecedingElementStatus = 8 -- 8 = failed action
									BEGIN
										SET @iTemp3 = CASE
												WHEN @iPrecedingElementStatus = 3 THEN 0
												ELSE 1
											END;
		
										SELECT @iCount = COUNT(*)
										FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
										WHERE ID = @iElementID;
		
										IF @iCount > 0 SET @fValidElements = 1;
									END
								END
								ELSE
								BEGIN
									-- Preceding element must have status 3 (3 =Completed)
									IF @iPrecedingElementStatus = 3 SET @fValidElements = 1;
								END
							END
						END
		
						FETCH NEXT FROM precedingElementsCursor INTO  @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow;
					END
					CLOSE precedingElementsCursor;
					DEALLOCATE precedingElementsCursor;
		
					-- If all preceding steps have been completed submit the Or step.
					IF @fValidElements > 0 
					BEGIN
						-- Cancel any preceding steps that are not completed as they are no longer required.
						EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @iInstanceID, @iElementID;
		
						SET @iAction = 1;
					END
				END
		
				IF @iAction = 1
				BEGIN
					EXEC [dbo].[spASRSubmitWorkflowStep] @iInstanceID, @iElementID, '''', @sForms OUTPUT, @fSaveForLater OUTPUT;
				END
		
				IF @iAction = 2
				BEGIN
					UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
					SET status = 2
					WHERE id = @iStepID;
				END
		
				FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID;
			END
		
			CLOSE stepsCursor;
			DEALLOCATE stepsCursor;
		
			DECLARE timeoutCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT 
				WIS.instanceID,
				WE.ID,
				WIS.ID
			FROM ASRSysWorkflowInstanceSteps WIS
			INNER JOIN ASRSysWorkflowElements WE ON WIS.elementID = WE.ID
				AND WE.type = 2 -- WebForm
			WHERE WIS.status = 2 -- Pending user action
				AND isnull(WE.timeoutFrequency,0) > 0
				AND CASE 
					WHEN WE.timeoutPeriod = 0 THEN datediff(minute, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 1 THEN datediff(Hour, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 2 AND WE.timeoutExcludeWeekend = 1 THEN 
						datediff(day, dbo.udfASRAddWeekdays(WIS.activationDateTime,WE.timeoutFrequency), getDate())
					WHEN WE.timeoutPeriod = 2 THEN datediff(day, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 3 THEN datediff(week, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 4 THEN datediff(month, WIS.activationDateTime, getDate())
					WHEN WE.timeoutPeriod = 5 THEN datediff(year, WIS.activationDateTime, getDate())
					ELSE 0
				END >= WE.timeoutFrequency;
		
			OPEN timeoutCursor;
			FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID;
			WHILE (@@fetch_status = 0)
			BEGIN
				-- Set the step status to be Timeout
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 6, -- Timeout
					ASRSysWorkflowInstanceSteps.timeoutCount = isnull(ASRSysWorkflowInstanceSteps.timeoutCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID;
		
				-- Activate the succeeding elements on the Timeout flow
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 1,
					ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionDateTime = null
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID IN 
						(SELECT id 
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 1))
					AND (ASRSysWorkflowInstanceSteps.status = 0
						OR ASRSysWorkflowInstanceSteps.status = 3
						OR ASRSysWorkflowInstanceSteps.status = 4
						OR ASRSysWorkflowInstanceSteps.status = 6
						OR ASRSysWorkflowInstanceSteps.status = 8);
					
				-- Set activated Web Forms to be ''pending'' (to be done by the user)
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 2
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 2);
					
				-- Set activated Terminators to be ''completed''
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 1);
					
				-- Count how many terminators have completed. ie. if the workflow has completed.
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type = 1;
										
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @iInstanceID;
					
					-- NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					-- is performed by a DELETE trigger on the ASRSysWorkflowInstances table.
				END
		
				FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID;
			END
		
			CLOSE timeoutCursor;
			DEALLOCATE timeoutCursor;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRDelegateWorkflowEmail
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRDelegateWorkflowEmail]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRDelegateWorkflowEmail];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRDelegateWorkflowEmail]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRDelegateWorkflowEmail] 
		(
			@psTo						varchar(MAX),
			@psCopyTo					varchar(MAX),
			@psMessage					varchar(MAX),
			@psMessage_HypertextLinks	varchar(MAX),
			@piStepID					integer,
			@psEmailSubject				varchar(MAX)
		)
		AS
		BEGIN
			DECLARE
				@sTo				varchar(MAX),
				@sAddress			varchar(MAX),
				@iInstanceID		integer,
				@curRecipients		cursor,
				@sEmailAddress		varchar(MAX),
				@fDelegated			bit,
				@sDelegatedTo		varchar(MAX),
				@fIsDelegate		bit;
		
			SET @psMessage = isnull(@psMessage, '''');
			SET @psMessage_HypertextLinks = isnull(@psMessage_HypertextLinks, '''');
			IF (len(ltrim(rtrim(@psTo))) = 0) RETURN;
		
			-- Get the instanceID of the given step
			SELECT @iInstanceID = instanceID
			FROM ASRSysWorkflowInstanceSteps
			WHERE ID = @piStepID;
				
		    DECLARE @recipients TABLE (
				emailAddress	varchar(MAX),
				delegated		bit,
				delegatedTo		varchar(MAX),
				isDelegate		bit
		    )
		
			exec [dbo].[spASRGetWorkflowDelegates] 
				@psTo, 
				@piStepID, 
				@curRecipients output;
				
			FETCH NEXT FROM @curRecipients INTO 
					@sEmailAddress,
					@fDelegated,
					@sDelegatedTo,
					@fIsDelegate
			WHILE (@@fetch_status = 0)
			BEGIN
				INSERT INTO @recipients
					(emailAddress,
					delegated,
					delegatedTo,
					isDelegate)
				VALUES (
					@sEmailAddress,
					@fDelegated,
					@sDelegatedTo,
					@fIsDelegate
				);
				
				FETCH NEXT FROM @curRecipients INTO 
						@sEmailAddress,
						@fDelegated,
						@sDelegatedTo,
						@fIsDelegate;
			END
			CLOSE @curRecipients;
			DEALLOCATE @curRecipients;
		
			-- Clear out the delegation record for the current step
			DELETE FROM [dbo].[ASRSysWorkflowStepDelegation]
			WHERE stepID = @piStepID;
		
			INSERT INTO [dbo].[ASRSysWorkflowStepDelegation] (delegateEmail, stepID)
			SELECT DISTINCT emailAddress, @piStepID
			FROM @recipients
			WHERE isDelegate = 1;
		
			SET @sTo = '''';
			
			DECLARE toCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT DISTINCT ltrim(rtrim(emailAddress))
			FROM @recipients
			WHERE len(ltrim(rtrim(emailAddress))) > 0;
		
			OPEN toCursor;
			FETCH NEXT FROM toCursor INTO @sAddress;
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @sTo = @sTo
					+ CASE 
						WHEN len(ltrim(rtrim(@sTo))) > 0 THEN '';''
						ELSE ''''
					END 
					+ @sAddress;
		
				FETCH NEXT FROM toCursor INTO @sAddress;
			END
			CLOSE toCursor;
			DEALLOCATE toCursor;
		
			IF len(@sTo) > 0
			BEGIN
				INSERT [dbo].[ASRSysEmailQueue](
					RecordDesc,
					ColumnValue, 
					DateDue, 
					UserName, 
					[Immediate],
					RecalculateRecordDesc, 
					RepTo,
					MsgText,
					WorkflowInstanceID, 
					[Subject])
				VALUES ('''',
					'''',
					getdate(),
					''HR Pro Workflow'',
					1,
					0, 
					@sTo,
					@psMessage + @psMessage_HypertextLinks,
					@iInstanceID,
					@psEmailSubject);
			END
		
			IF (len(@psCopyTo) > 0) AND (len(@psMessage) > 0)
			BEGIN
				INSERT ASRSysEmailQueue(
					RecordDesc,
					ColumnValue, 
					DateDue, 
					UserName, 
					[Immediate],
					RecalculateRecordDesc, 
					RepTo,
					MsgText,
					WorkflowInstanceID, 
					[Subject])
				VALUES ('''',
					'''',
					getdate(),
					''HR Pro Workflow'',
					1,
					0, 
					@psCopyTo,
					''You have been copied in on the following HR Pro Workflow email with recipients:'' + CHAR(13)
						+ CHAR(9) + @sTo + CHAR(13)	+ CHAR(13)
						+ @psMessage,
					@iInstanceID,
					@psEmailSubject);
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASREmailBatch
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASREmailBatch]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASREmailBatch];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASREmailBatch]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASREmailBatch]
		AS
		BEGIN

			DECLARE @QueueID	integer,
				@LinkID			integer,
				@RecordID		integer,
				@ColumnID		integer,
				@ColumnValue	integer,
				@RecDescID		integer,
				@RecDesc		nvarchar(MAX),
				@sSQL			nvarchar(MAX),
				@EmailDate		datetime,
				@hResult		integer,
				@blnEnabled		integer;

			SELECT @blnEnabled = [SettingValue] FROM [dbo].[ASRSysSystemSettings]
				WHERE [Section] = ''email'' and [SettingKey] = ''overnight enabled'';

			IF @blnEnabled = 0
			BEGIN
				RETURN
			END

			-- Clear Servers Inbox
			-- Doing this just before sending messages means that any failure return messages will
			-- stay in the servers inbox until this sp is run again - could be useful for support ?

			-- DECLARE @message_id varchar(255)
			-- EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
			-- WHILE not @message_ID is null
			-- BEGIN
			--	EXEC master.dbo.xp_deletemail @message_id
			--	SET @message_id = null
			--	EXEC master.dbo.xp_findnextmsg @msg_id = @message_id output
			-- END


			/* Purge email queue */
			EXEC sp_ASRPurgeRecords ''EMAIL'', ''ASRSysEmailQueue'', ''DateDue'';

			/* Send all emails waiting to be sent regardless of username */
			EXEC spASREmailImmediate '''';

		END';

	EXECUTE sp_executeSQL @sSPCode;



	----------------------------------------------------------------------
	-- spASREmailRebuild
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASREmailRebuild]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASREmailRebuild];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASREmailRebuild]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASREmailRebuild]
		AS
		BEGIN	
			/* Refresh all calculated columns in the database. */
			DECLARE @sTableName 	varchar(255),
				@iTableID			integer,
				@sSQL				nvarchar(MAX),
				@sColumnName		varchar(255);
		
			
			/* Get a cursor of the tables in the database. */
			DECLARE curTables CURSOR FOR
				SELECT tableName, tableID
				FROM ASRSysTables
			OPEN curTables;
		
			DELETE FROM AsrSysEmailQueue WHERE DateSent Is Null AND [Immediate] = 0;
		
			/* Loop through the tables in the database. */
			FETCH NEXT FROM curTables INTO @sTableName, @iTableID;
			WHILE @@fetch_status <> -1
			BEGIN
				/* Get a cursor of the records in the current table.  */
				/* Call the diary trigger for that table and record  */
				SET @sSQL = ''DECLARE @iCurrentID	int,
								@sSQL		nvarchar(MAX);
							
							IF EXISTS (SELECT * FROM sysobjects
							WHERE id = object_id(''''spASREmailRebuild_'' + LTrim(Str(@iTableID)) + '''''') 
								AND sysstat & 0xf = 4)
							BEGIN
								DECLARE curRecords CURSOR FOR
								SELECT id
								FROM '' + @sTableName + '';
				
								OPEN curRecords;
				
								FETCH NEXT FROM curRecords INTO @iCurrentID;
								WHILE @@fetch_status <> -1
								BEGIN
									PRINT ''''ID : '''' + Str(@iCurrentID);
									SET @sSQL = ''''EXEC spASREmailRebuild_'' + LTrim(Str(@iTableID)) 
										+ '' '''' + convert(varchar(100), @iCurrentID) + '''''''';
									EXECUTE sp_executeSQL @sSQL;
				
									FETCH NEXT FROM curRecords INTO @iCurrentID;
								END
								CLOSE curRecords;
								DEALLOCATE curRecords;
							END'';
				 EXECUTE sp_executeSQL @sSQL;
		
				/* Move onto the next table in the database. */ 
				FETCH NEXT FROM curTables INTO @sTableName, @iTableID;
			END
		
			CLOSE curTables;
			DEALLOCATE curTables;
		
			EXEC [dbo].spASREmailImmediate '''';
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

		----------------------------------------------------------------------
	-- spASRGetCurrentUsersAppName
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersAppName]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersAppName]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersAppName]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetCurrentUsersAppName]
		(
			@psAppName		varchar(MAX) OUTPUT,
			@psUserName		varchar(MAX)
		)
		AS
		BEGIN
		
		    IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
		    BEGIN
		        EXEC sp_ASRIntCheckPolls;
		    END
		
		    SELECT TOP 1 @psAppName = rtrim(p.program_name)
		    FROM master..sysprocesses p
		    WHERE p.program_name LIKE ''HR Pro%''
				AND		p.program_name NOT LIKE ''HR Pro Workflow%''
		    AND		p.program_name NOT LIKE ''HR Pro Outlook%''
		    AND		p.program_name NOT LIKE ''HR Pro Server.Net%''
				AND		p.program_name NOT LIKE ''HR Pro Intranet Embedding%''
				AND		p.loginame = @psUsername
		    GROUP BY p.hostname
		           , p.loginame
		           , p.program_name
		           , p.hostprocess
		    ORDER BY p.loginame;
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetCurrentUsersCountInApp
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersCountInApp]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp]
		(
			@piCount integer OUTPUT
		)
		AS
		BEGIN

			SET NOCOUNT ON;

			DECLARE @sSQLVersion	integer;
			DECLARE @Mode			smallint;

			IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
			BEGIN
				EXEC sp_ASRIntCheckPolls;
			END

			SELECT @sSQLVersion = dbo.udfASRSQLVersion();
			SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = ''ProcessAccount'' AND [SettingKey] = ''Mode'';
			IF @@ROWCOUNT = 0 SET @Mode = 0;

			IF ((@Mode = 1 OR @Mode = 2) AND @sSQLVersion > 8) AND (NOT IS_SRVROLEMEMBER(''sysadmin'') = 1)
			BEGIN
				SELECT @piCount = dbo.[udfASRNetCountCurrentUsersInApp](APP_NAME());
			END
			ELSE
			BEGIN

				SELECT @piCount = COUNT(p.Program_Name)
				FROM     master..sysprocesses p
				JOIN     master..sysdatabases d
				  ON     d.dbid = p.dbid
				WHERE    p.program_name = APP_NAME()
				  AND    d.name = db_name()
				GROUP BY p.program_name;
			END

		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetCurrentUsersCountOnServer
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersCountOnServer]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersCountOnServer]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersCountOnServer]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetCurrentUsersCountOnServer]
		(
			@iLoginCount	integer OUTPUT,
			@psLoginName	varchar(MAX)
		)
		AS
		BEGIN
		
			DECLARE @sSQLVersion int;
			DECLARE @Mode smallint;
		
			IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
			BEGIN
				EXEC sp_ASRIntCheckPolls;
			END
		
			SELECT @sSQLVersion = dbo.udfASRSQLVersion();
			SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = ''ProcessAccount'' AND [SettingKey] = ''Mode'';
			IF @@ROWCOUNT = 0 SET @Mode = 0

			IF ((@Mode = 1 OR @Mode = 2) AND @sSQLVersion > 8) AND (NOT IS_SRVROLEMEMBER(''sysadmin'') = 1)		
			BEGIN
				SELECT @iLoginCount = dbo.[udfASRNetCountCurrentLogins](@psLoginName);
			END
			ELSE
			BEGIN
		
				SELECT @iLoginCount = COUNT(*)
				FROM master..sysprocesses p
				WHERE p.program_name LIKE ''HR Pro%''
				AND		p.program_name NOT LIKE ''HR Pro Workflow%''
		    AND		p.program_name NOT LIKE ''HR Pro Outlook%''
		    AND		p.program_name NOT LIKE ''HR Pro Server.Net%''
		    AND		p.program_name NOT LIKE ''HR Pro Intranet Embedding%''
		
				    AND p.loginame = @psLoginName;
			END
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetCurrentUsersInGroups
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersInGroups]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersInGroups];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersInGroups]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE spASRGetCurrentUsersInGroups 
		(
			@psGroupNames VARCHAR(MAX)
		)
		AS
		BEGIN
			SET NOCOUNT ON
		
			CREATE TABLE #tblCurrentUsers				
			(
				hostname varchar(256)
				,loginame varchar(256)
				,program_name varchar(256)
				,hostprocess varchar(20)
				,sid binary(86)
				,login_time datetime
				,spid int
				,uid int
			)
			INSERT INTO #tblCurrentUsers
				EXEC spASRGetCurrentUsers
		
			DECLARE @tblGroups TABLE
			(
				groupname varchar(256) collate database_default 
			)
		
			DECLARE @IN			varchar(MAX), 
					@INGroup	varchar(MAX),
					@Pos		integer;
		
			SET @psGroupNames = LTRIM(RTRIM(@psGroupNames))+ '',''
			SET @Pos = CHARINDEX('','', @psGroupNames, 1)
			SET @IN = ''''
		
			IF REPLACE(@psGroupNames, '','', '''') <> ''''
			BEGIN
				WHILE @Pos > 0
				BEGIN
					SET @INGroup = LTRIM(RTRIM(LEFT(@psGroupNames, @Pos - 1)))
					IF @INGroup <> ''''
					BEGIN
						INSERT INTO @tblGroups VALUES (@INGroup)
					END
					SET @psGroupNames = RIGHT(@psGroupNames, LEN(@psGroupNames) - @Pos)
					SET @Pos = CHARINDEX('','', @psGroupNames, 1)
				END
			END
		
			SELECT [#tblCurrentUsers].[loginame]
			FROM [#tblCurrentUsers] 
				JOIN [ASRSysUserGroups] ON [#tblCurrentUsers].[loginame] = [ASRSysUserGroups].[UserName] collate database_default 
			WHERE [ASRSysUserGroups].[UserGroup] IN 
				(
					SELECT [groupName] FROM @tblGroups
				)
		
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetCurrentUsersInWindowsGroups
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersInWindowsGroups]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersInWindowsGroups];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersInWindowsGroups]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetCurrentUsersInWindowsGroups]
		(
			@psGroupNames VARCHAR(MAX)
		)
		AS
		BEGIN
			SET NOCOUNT ON;
		
			CREATE TABLE #tblCurrentUsers				
				(
					hostname varchar(256)
					,loginame varchar(256)
					,program_name varchar(256)
					,hostprocess varchar(20)
					,sid binary(86)
					,login_time datetime
					,spid int
					,uid int
				)
				
			INSERT INTO #tblCurrentUsers
				EXEC spASRGetCurrentUsers
		
			CREATE TABLE #tblUsersInGroup
				(
				loginame varchar(256)
				,groupname varchar(256)
				)		
		
			DECLARE @tblGroups TABLE
				(
				groupname varchar(256)
				)
		
			DECLARE @iUserInGroup	integer,
					@loginame		varchar(256),
					@IN				varchar(MAX), 
					@INGroup		varchar(MAX),
					@Pos			int;
		
			SET @psGroupNames = LTRIM(RTRIM(@psGroupNames))+ '',''
			SET @Pos = CHARINDEX('','', @psGroupNames, 1)
			SET @IN = ''''
		
			IF REPLACE(@psGroupNames, '','', '''') <> ''''
			BEGIN
				WHILE @Pos > 0
				BEGIN
					SET @INGroup = LTRIM(RTRIM(LEFT(@psGroupNames, @Pos - 1)))
					IF @INGroup <> ''''
					BEGIN
						INSERT INTO @tblGroups VALUES (@INGroup)
					END
					SET @psGroupNames = RIGHT(@psGroupNames, LEN(@psGroupNames) - @Pos)
					SET @Pos = CHARINDEX('','', @psGroupNames, 1)
				END
			END
		
			DECLARE CurrentUsersCursor CURSOR LOCAL FAST_FORWARD READ_ONLY FOR 
			SELECT loginame FROM #tblCurrentUsers
			OPEN CurrentUsersCursor
			FETCH NEXT FROM CurrentUsersCursor INTO @loginame
		
			IF @@FETCH_STATUS <> 0 
				RETURN
		
			WHILE @@FETCH_STATUS = 0
			BEGIN
				SET @loginame = LTRIM(RTRIM(@loginame))
		
				INSERT INTO #tblUsersInGroup 
					EXEC spASRGroupsUserIsMemberOf @loginame
		
				FETCH NEXT FROM CurrentUsersCursor INTO @loginame
			END
			CLOSE CurrentUsersCursor
			DEALLOCATE CurrentUsersCursor
		
			DROP TABLE #tblCurrentUsers
		
			/* Return a recordset of the users to log out */
			SELECT loginame
			FROM #tblUsersInGroup 
			WHERE groupname IN (SELECT DISTINCT groupname collate database_default FROM @tblGroups)
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetDomains
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetDomains]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetDomains];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetDomains]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetDomains]
			(@DomainString varchar(MAX) OUTPUT)
		AS
		BEGIN
		
			SELECT @DomainString = dbo.udfASRNetGetDomains();
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetParentDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetParentDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetParentDetails];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetParentDetails]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetParentDetails]
		(
			@piBaseTableID		integer,
			@piBaseRecordID		integer,
			@piParent1TableID	integer	OUTPUT,
			@piParent1RecordID	integer	OUTPUT,
			@piParent2TableID	integer	OUTPUT,
			@piParent2RecordID	integer	OUTPUT
		)
		AS
		BEGIN
			-- Return the parent table IDs and related record IDs for the given table and record.
			-- Return 0 if no table/record exists.
			DECLARE
				@sSQL		nvarchar(MAX),
				@sParam		nvarchar(500),
				@sTableName	nvarchar(255);
		
			SET @piParent1TableID = 0;
			SET @piParent1RecordID = 0;
			SET @piParent2TableID = 0;
			SET @piParent2RecordID = 0;
		
			SELECT @sTableName = tableName
			FROM [dbo].[ASRSysTables]
			WHERE tableID = @piBaseTableID;
		
			SELECT TOP 1 @piParent1TableID = isnull(parentID, 0)
			FROM [dbo].[ASRSysRelations]
			WHERE childID = @piBaseTableID
			ORDER BY parentID ASC;
		
			SELECT TOP 1 @piParent2TableID = isnull(parentID, 0)
			FROM [dbo].[ASRSysRelations] 
			WHERE childID = @piBaseTableID
				AND parentID <> @piParent1TableID
			ORDER BY parentID ASC;
		
			IF (LEN(@sTableName) > 0) AND (@piBaseRecordID > 0)
			BEGIN
				IF (@piParent1TableID > 0)
				BEGIN
					SET @sSQL = ''SELECT @piParent1RecordID = isnull(ID_'' + convert(nvarchar(100), @piParent1TableID) + '',0)''
						+ '' FROM '' + @sTableName
						+ '' WHERE ID = '' + convert(varchar(100), @piBaseRecordID);
					SET @sParam = N''@piParent1RecordID integer OUTPUT'';
					EXEC sp_executesql @sSQL, @sParam, @piParent1RecordID OUTPUT;
				END
		
				IF @piParent2TableID > 0 
				BEGIN
					SET @sSQL = ''SELECT @piParent2RecordID = isnull(ID_'' + convert(nvarchar(100), @piParent2TableID) + '',0)''
						+ '' FROM '' + @sTableName
						+ '' WHERE ID = '' + convert(varchar(100), @piBaseRecordID);
					SET @sParam = N''@piParent2RecordID integer OUTPUT'';
					EXEC sp_executesql @sSQL, @sParam, @piParent2RecordID OUTPUT;
				END		
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetSetting
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetSetting]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetSetting];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetSetting]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetSetting] (
			@psSection		varchar(25),
			@psKey			varchar(255),
			@psDefault		varchar(MAX),
			@pfUserSetting	bit,
			@psResult		varchar(MAX) OUTPUT
		)
		AS
		BEGIN
			/* Return the required user or system setting. */
			DECLARE	@iCount	integer;
		
			IF @pfUserSetting = 1
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM [dbo].[ASRSysUserSettings]
				WHERE userName = SYSTEM_USER
					AND section = @psSection		
					AND settingKey = @psKey;
		
				SELECT @psResult = settingValue 
				FROM [dbo].[ASRSysUserSettings]
				WHERE userName = SYSTEM_USER
					AND section = @psSection		
					AND settingKey = @psKey;
			END
			ELSE
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM [dbo].[ASRSysSystemSettings]
				WHERE section = @psSection		
					AND settingKey = @psKey;
		
				SELECT @psResult = settingValue 
				FROM [dbo].[ASRSysSystemSettings]
				WHERE section = @psSection		
					AND settingKey = @psKey;
			END
		
			IF @iCount = 0
			BEGIN
				SET @psResult = @psDefault;	
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetStoredDataActionDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetStoredDataActionDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetStoredDataActionDetails];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetStoredDataActionDetails]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetStoredDataActionDetails]
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@psSQL				varchar(MAX)	OUTPUT, 
			@piDataTableID		integer			OUTPUT,
			@psTableName		varchar(255)	OUTPUT,
			@piDataAction		integer			OUTPUT, 
			@piRecordID			integer			OUTPUT
		)
		AS
		BEGIN
			DECLARE 
				@iPersonnelTableID			integer,
				@iInitiatorID				integer,
				@iDataRecord				integer,
				@sIDColumnName				varchar(MAX),
				@iColumnID					integer, 
				@sColumnName				varchar(MAX), 
				@iColumnDataType			integer, 
				@sColumnList				varchar(MAX),
				@sValueList					varchar(MAX),
				@sValue						varchar(MAX),
				@sRecSelWebFormIdentifier	varchar(MAX),
				@sRecSelIdentifier			varchar(MAX),
				@iTempTableID				integer,
				@iSecondaryDataRecord		integer,
				@sSecondaryRecSelWebFormIdentifier	varchar(MAX),
				@sSecondaryRecSelIdentifier	varchar(MAX),
				@sSecondaryIDColumnName		varchar(MAX),
				@iSecondaryRecordID			integer,
				@iElementType				integer,
				@iWorkflowID				integer,
				@iID						integer,
				@sWFFormIdentifier			varchar(MAX),
				@sWFValueIdentifier			varchar(MAX),
				@iDBColumnID				integer,
				@iDBRecord					integer,
				@sSQL						nvarchar(MAX),
				@sParam						nvarchar(MAX),
				@sDBColumnName				nvarchar(MAX),
				@sDBTableName				nvarchar(MAX),
				@iRecordID					integer,
				@sDBValue					varchar(MAX),
				@iDataType					integer, 
				@iValueType					integer, 
				@iSDColumnID				integer,
				@fValidRecordID				bit,
				@iBaseTableID				integer,
				@iBaseRecordID				integer,
				@iRequiredTableID			integer,
				@iRequiredRecordID			integer,
				@iDataRecordTableID			integer,
				@iSecondaryDataRecordTableID	integer,
				@iParent1TableID			integer,
				@iParent1RecordID			integer,
				@iParent2TableID			integer,
				@iParent2RecordID			integer,
				@iInitParent1TableID		integer,
				@iInitParent1RecordID		integer,
				@iInitParent2TableID		integer,
				@iInitParent2RecordID		integer,
				@iEmailID					integer,
				@iType						integer,
				@fDeletedValue				bit,
				@iTempElementID				integer,
				@iCount						integer,
				@iResultType				integer,
				@sResult					varchar(MAX),
				@fResult					bit,
				@dtResult					datetime,
				@fltResult					float,
				@iCalcID					integer,
				@iSize						integer,
				@iDecimals					integer,
				@iTriggerTableID			integer;
					
			SET @psSQL = '''';
			SET @piRecordID = 0;
		
			SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_PERSONNEL''
				AND parameterKey = ''Param_TablePersonnel'';
		
			IF @iPersonnelTableID = 0
			BEGIN
				SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_TablePersonnel'';
			END
		
			SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
				@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
				@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
				@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
				@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
			FROM ASRSysWorkflowInstances
			WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
		
			SELECT @piDataAction = dataAction,
				@piDataTableID = dataTableID,
				@iDataRecord = dataRecord,
				@sRecSelWebFormIdentifier = recSelWebFormIdentifier,
				@sRecSelIdentifier = recSelIdentifier,
				@iSecondaryDataRecord = secondaryDataRecord,
				@sSecondaryRecSelWebFormIdentifier = secondaryRecSelWebFormIdentifier,
				@sSecondaryRecSelIdentifier = secondaryRecSelIdentifier,
				@iDataRecordTableID = dataRecordTable,
				@iSecondaryDataRecordTableID = secondaryDataRecordTable,
				@iWorkflowID = workflowID,
				@iTriggerTableID = ASRSysWorkflows.baseTable
			FROM ASRSysWorkflowElements
			INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID
			WHERE ASRSysWorkflowElements.ID = @piElementID;
		
			SELECT @psTableName = tableName
			FROM ASRSysTables
			WHERE tableID = @piDataTableID;
		
			IF @iDataRecord = 0 -- 0 = Initiator''s record
			BEGIN
				EXEC [dbo].[spASRWorkflowAscendantRecordID]
					@iPersonnelTableID,
					@iInitiatorID,
					@iInitParent1TableID,
					@iInitParent1RecordID,
					@iInitParent2TableID,
					@iInitParent2RecordID,
					@iDataRecordTableID,
					@piRecordID	OUTPUT;
		
				IF @piDataTableID = @iDataRecordTableID
				BEGIN
					SET @sIDColumnName = ''ID'';
				END
				ELSE
				BEGIN
					SET @sIDColumnName = ''ID_'' + convert(varchar(255), @iDataRecordTableID);
				END
			END
		
			IF @iDataRecord = 4 -- 4 = Triggered record
			BEGIN
				EXEC [dbo].[spASRWorkflowAscendantRecordID]
					@iTriggerTableID,
					@iInitiatorID,
					@iInitParent1TableID,
					@iInitParent1RecordID,
					@iInitParent2TableID,
					@iInitParent2RecordID,
					@iDataRecordTableID,
					@piRecordID	OUTPUT;
		
				IF @piDataTableID = @iDataRecordTableID
				BEGIN
					SET @sIDColumnName = ''ID'';
				END
				ELSE
				BEGIN
					SET @sIDColumnName = ''ID_'' + convert(varchar(255), @iDataRecordTableID);
				END
			END
		
			IF @iDataRecord = 1 -- 1 = Identified record
			BEGIN
				SELECT @iElementType = ASRSysWorkflowElements.type
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)));
				
				IF @iElementType = 2
				BEGIN
					 -- WebForm
					SELECT @sValue = ISNULL(IV.value, ''0''),
						@iTempTableID = EI.tableID,
						@iParent1TableID = IV.parent1TableID,
						@iParent1RecordID = IV.parent1RecordID,
						@iParent2TableID = IV.parent2TableID,
						@iParent2RecordID = IV.parent2RecordID
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
					INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
					WHERE IV.instanceID = @piInstanceID
						AND IV.identifier = @sRecSelIdentifier
						AND Es.identifier = @sRecSelWebFormIdentifier
						AND Es.workflowID = @iWorkflowID
						AND IV.elementID = Es.ID;
				END
				ELSE
				BEGIN
					-- StoredData
					SELECT @sValue = ISNULL(IV.value, ''0''),
						@iTempTableID = Es.dataTableID,
						@iParent1TableID = IV.parent1TableID,
						@iParent1RecordID = IV.parent1RecordID,
						@iParent2TableID = IV.parent2TableID,
						@iParent2RecordID = IV.parent2RecordID
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
						AND IV.identifier = Es.identifier
						AND Es.workflowID = @iWorkflowID
						AND Es.identifier = @sRecSelWebFormIdentifier
					WHERE IV.instanceID = @piInstanceID;
				END
		
				SET @piRecordID = 
					CASE
						WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
						ELSE 0
					END;
			
				SET @iBaseTableID = @iTempTableID;
				SET @iBaseRecordID = @piRecordID;
				EXEC [dbo].[spASRWorkflowAscendantRecordID]
					@iBaseTableID,
					@iBaseRecordID,
					@iParent1TableID,
					@iParent1RecordID,
					@iParent2TableID,
					@iParent2RecordID,
					@iDataRecordTableID,
					@piRecordID	OUTPUT;
		
				IF @piDataTableID = @iDataRecordTableID
				BEGIN
					SET @sIDColumnName = ''ID'';
				END
				ELSE
				BEGIN
					SET @sIDColumnName = ''ID_'' + convert(varchar(255), @iDataRecordTableID);
				END
			END
		
			SET @fValidRecordID = 1
			IF (@iDataRecord = 0) OR (@iDataRecord = 1) OR (@iDataRecord = 4)
			BEGIN
				EXEC [dbo].[spASRWorkflowValidTableRecord]
					@iDataRecordTableID,
					@piRecordID,
					@fValidRecordID	OUTPUT;
		
				IF @fValidRecordID = 0
				BEGIN
					-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
					EXEC [dbo].[spASRWorkflowActionFailed]
						@piInstanceID, 
						@piElementID, 
						''Stored Data primary record has been deleted or not selected.'';
		
					SET @psSQL = '''';
					RETURN;
				END
			END
		
			IF @piDataAction = 0 -- Insert
			BEGIN
				IF @iSecondaryDataRecord = 0 -- 0 = Initiator''s record
				BEGIN
					EXEC [dbo].[spASRWorkflowAscendantRecordID]
						@iPersonnelTableID,
						@iInitiatorID,
						@iInitParent1TableID,
						@iInitParent1RecordID,
						@iInitParent2TableID,
						@iInitParent2RecordID,
						@iSecondaryDataRecordTableID,
						@iSecondaryRecordID	OUTPUT;
			
					IF @piDataTableID = @iSecondaryDataRecordTableID
					BEGIN
						SET @sSecondaryIDColumnName = ''ID'';
					END
					ELSE
					BEGIN
						SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(255), @iSecondaryDataRecordTableID);
					END
				END
				
				IF @iSecondaryDataRecord = 4 -- 4 = Triggered record
				BEGIN
					EXEC [dbo].[spASRWorkflowAscendantRecordID]
						@iTriggerTableID,
						@iInitiatorID,
						@iInitParent1TableID,
						@iInitParent1RecordID,
						@iInitParent2TableID,
						@iInitParent2RecordID,
						@iSecondaryDataRecordTableID,
						@iSecondaryRecordID	OUTPUT;
			
					IF @piDataTableID = @iSecondaryDataRecordTableID
					BEGIN
						SET @sSecondaryIDColumnName = ''ID'';
					END
					ELSE
					BEGIN
						SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(255), @iSecondaryDataRecordTableID);
					END
				END
		
				IF @iSecondaryDataRecord = 1 -- 1 = Previous record selector''s record
				BEGIN
					SELECT @iElementType = ASRSysWorkflowElements.type
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
						AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sSecondaryRecSelWebFormIdentifier)));
			
					IF @iElementType = 2
					BEGIN
						 -- WebForm
						SELECT @sValue = ISNULL(IV.value, ''0''),
							@iTempTableID = EI.tableID,
							@iParent1TableID = IV.parent1TableID,
							@iParent1RecordID = IV.parent1RecordID,
							@iParent2TableID = IV.parent2TableID,
							@iParent2RecordID = IV.parent2RecordID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
						INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
						WHERE IV.instanceID = @piInstanceID
							AND IV.identifier = @sSecondaryRecSelIdentifier
							AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
							AND Es.workflowID = @iWorkflowID
							AND IV.elementID = Es.ID;
					END
					ELSE
					BEGIN
						-- StoredData
						SELECT @sValue = ISNULL(IV.value, ''0''),
							@iTempTableID = Es.dataTableID,
							@iParent1TableID = IV.parent1TableID,
							@iParent1RecordID = IV.parent1RecordID,
							@iParent2TableID = IV.parent2TableID,
							@iParent2RecordID = IV.parent2RecordID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
							AND IV.identifier = Es.identifier
							AND Es.workflowID = @iWorkflowID
							AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
						WHERE IV.instanceID = @piInstanceID;
					END
		
					SET @iSecondaryRecordID = 
						CASE
							WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
							ELSE 0
						END;
					
					SET @iBaseTableID = @iTempTableID;
					SET @iBaseRecordID = @iSecondaryRecordID;
					EXEC [dbo].[spASRWorkflowAscendantRecordID]
						@iBaseTableID,
						@iBaseRecordID,
						@iParent1TableID,
						@iParent1RecordID,
						@iParent2TableID,
						@iParent2RecordID,
						@iSecondaryDataRecordTableID,
						@iSecondaryRecordID	OUTPUT;
		
					IF @piDataTableID = @iSecondaryDataRecordTableID
					BEGIN
						SET @sSecondaryIDColumnName = ''ID'';
					END
					ELSE
					BEGIN
						SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(255), @iSecondaryDataRecordTableID);
					END
				END
		
				SET @fValidRecordID = 1;
				IF (@iSecondaryDataRecord = 0) OR (@iSecondaryDataRecord = 1) OR (@iSecondaryDataRecord = 4)
				BEGIN
					EXEC [dbo].[spASRWorkflowValidTableRecord]
						@iSecondaryDataRecordTableID,
						@iSecondaryRecordID,
						@fValidRecordID	OUTPUT;
		
					IF @fValidRecordID = 0
					BEGIN
						-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
						EXEC [dbo].[spASRWorkflowActionFailed] 
							@piInstanceID, 
							@piElementID, 
							''Stored Data secondary record has been deleted or not selected.'';
		
						SET @psSQL = '''';
						RETURN;
					END
				END
			END
		
			IF @piDataAction = 0 OR @piDataAction = 1
			BEGIN
				/* INSERT or UPDATE. */
				SET @sColumnList = '''';
				SET @sValueList = '''';
		
				DECLARE @dbValues TABLE (
					ID integer, 
					wfFormIdentifier varchar(1000),
					wfValueIdentifier varchar(1000),
					dbColumnID int,
					dbRecord int,
					value varchar(MAX));
		
				INSERT INTO @dbValues (ID, 
					wfFormIdentifier,
					wfValueIdentifier,
					dbColumnID,
					dbRecord,
					value) 
				SELECT EC.ID,
					EC.wfformidentifier,
					EC.wfvalueidentifier,
					EC.dbcolumnid,
					EC.dbrecord, 
					''''
				FROM ASRSysWorkflowElementColumns EC
				WHERE EC.elementID = @piElementID
					AND EC.valueType = 2;
					
				DECLARE dbValuesCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ID,
					wfFormIdentifier,
					wfValueIdentifier,
					dbColumnID,
					dbRecord
				FROM @dbValues;
				OPEN dbValuesCursor;
				FETCH NEXT FROM dbValuesCursor INTO @iID,
					@sWFFormIdentifier,
					@sWFValueIdentifier,
					@iDBColumnID,
					@iDBRecord;
				WHILE (@@fetch_status = 0)
				BEGIN
					SET @fDeletedValue = 0;
		
					SELECT @sDBTableName = tbl.tableName,
						@iRequiredTableID = tbl.tableID, 
						@sDBColumnName = col.columnName,
						@iDataType = col.dataType
					FROM ASRSysColumns col
					INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
					WHERE col.columnID = @iDBColumnID;
		
					SET @sSQL = ''SELECT @sDBValue = ''
						+ CASE
							WHEN @iDataType = 12 THEN ''''
							WHEN @iDataType = 11 THEN ''convert(varchar(MAX),''
							ELSE ''convert(varchar(MAX),''
						END
						+ @sDBTableName + ''.'' + @sDBColumnName
						+ CASE
							WHEN @iDataType = 12 THEN ''''
							WHEN @iDataType = 11 THEN '', 101)''
							ELSE '')''
						END
						+ '' FROM '' + @sDBTableName 
						+ '' WHERE '' + @sDBTableName + ''.ID = '';
		
					SET @iRecordID = 0;
		
					IF @iDBRecord = 0
					BEGIN
						-- Initiator''s record
						SET @iRecordID = @iInitiatorID;
						SET @iParent1TableID = @iInitParent1TableID;
						SET @iParent1RecordID = @iInitParent1RecordID;
						SET @iParent2TableID = @iInitParent2TableID;
						SET @iParent2RecordID = @iInitParent2RecordID;
						SET @iBaseTableID = @iPersonnelTableID;
					END			
		
					IF @iDBRecord = 4
					BEGIN
						-- Trigger record
						SET @iRecordID = @iInitiatorID;
						SET @iParent1TableID = @iInitParent1TableID;
						SET @iParent1RecordID = @iInitParent1RecordID;
						SET @iParent2TableID = @iInitParent2TableID;
						SET @iParent2RecordID = @iInitParent2RecordID;
		
						SELECT @iBaseTableID = isnull(WF.baseTable, 0)
						FROM ASRSysWorkflows WF
						INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
							AND WFI.ID = @piInstanceID;
					END
					
					IF @iDBRecord = 1
					BEGIN
						-- Identified record
						SELECT @iElementType = ASRSysWorkflowElements.type, 
							@iTempElementID = ASRSysWorkflowElements.ID
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sWFFormIdentifier)));
		
						IF @iElementType = 2
						BEGIN
							 -- WebForm
							SELECT @sValue = ISNULL(IV.value, ''0''),
								@iBaseTableID = EI.tableID,
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
							INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
							WHERE IV.instanceID = @piInstanceID
								AND IV.identifier = @sWFValueIdentifier
								AND Es.identifier = @sWFFormIdentifier
								AND Es.workflowID = @iWorkflowID
								AND IV.elementID = Es.ID;
						END
						ELSE
						BEGIN
							-- StoredData
							SELECT @sValue = ISNULL(IV.value, ''0''),
								@iBaseTableID = isnull(Es.dataTableID, 0),
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
								AND IV.identifier = Es.identifier
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sWFFormIdentifier
							WHERE IV.instanceID = @piInstanceID;
						END
		
						SET @iRecordID = 
							CASE
								WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
								ELSE 0
							END;
					END
		
					SET @iBaseRecordID = @iRecordID;
		
					SET @fValidRecordID = 1;
					
					IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
					BEGIN
						SET @fValidRecordID = 0;
		
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
							@iBaseTableID,
							@iBaseRecordID,
							@iParent1TableID,
							@iParent1RecordID,
							@iParent2TableID,
							@iParent2RecordID,
							@iRequiredTableID,
							@iRequiredRecordID	OUTPUT;
		
						SET @iRecordID = @iRequiredRecordID;
		
						IF @iRecordID > 0 
						BEGIN
							EXEC [dbo].[spASRWorkflowValidTableRecord]
								@iRequiredTableID,
								@iRecordID,
								@fValidRecordID	OUTPUT;
						END
		
						IF @fValidRecordID = 0
						BEGIN
							IF @iDBRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
							BEGIN
								SELECT @iCount = COUNT(*)
								FROM ASRSysWorkflowQueueColumns QC
								INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
								WHERE WFQ.instanceID = @piInstanceID
									AND QC.columnID = @iDBColumnID;
		
								IF @iCount = 1
								BEGIN
									SELECT @sDBValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
									FROM ASRSysWorkflowQueueColumns QC
									INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
									WHERE WFQ.instanceID = @piInstanceID
										AND QC.columnID = @iDBColumnID;
		
									SET @fValidRecordID = 1;
									SET @fDeletedValue = 1;
								END
							END
							ELSE
							BEGIN
								IF @iDBRecord = 1
								BEGIN
									SELECT @iCount = COUNT(*)
									FROM ASRSysWorkflowInstanceValues IV
									WHERE IV.instanceID = @piInstanceID
										AND IV.columnID = @iDBColumnID
										AND IV.elementID = @iTempElementID;
		
									IF @iCount = 1
									BEGIN
										SELECT @sDBValue = rtrim(ltrim(isnull(IV.value , '''')))
										FROM ASRSysWorkflowInstanceValues IV
										WHERE IV.instanceID = @piInstanceID
											AND IV.columnID = @iDBColumnID
											AND IV.elementID = @iTempElementID;
		
										SET @fValidRecordID = 1;
										SET @fDeletedValue = 1;
									END
								END
							END
						END
		
						IF @fValidRecordID = 0
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC [dbo].[spASRWorkflowActionFailed]
								@piInstanceID, 
								@piElementID, 
								''Stored Data column database value record has been deleted or not selected.'';
		
							SET @psSQL = '''';
							RETURN;
						END
					END
		
					IF (@iDataType <> -3)
						AND (@iDataType <> -4)
					BEGIN
						IF @fDeletedValue = 0
						BEGIN
							SET @sSQL = @sSQL + convert(nvarchar(255), @iRecordID);
							SET @sParam = N''@sDBValue varchar(MAX) OUTPUT'';
							EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;
						END
		
						UPDATE @dbValues
						SET value = @sDBValue
						WHERE ID = @iID;
					END
					
					FETCH NEXT FROM dbValuesCursor INTO @iID,
						@sWFFormIdentifier,
						@sWFValueIdentifier,
						@iDBColumnID,
						@iDBRecord;
				END
				CLOSE dbValuesCursor;
				DEALLOCATE dbValuesCursor;
		
				DECLARE columnCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT EC.columnID,
					SC.columnName,
					SC.dataType,
					CASE
						WHEN EC.valueType = 0 THEN  -- Fixed Value
							CASE
								WHEN SC.dataType = -7 THEN
									CASE 
										WHEN UPPER(EC.value) = ''TRUE'' THEN ''1''
										ELSE ''0''
									END
								ELSE EC.value
							END
						WHEN EC.valueType = 1 THEN -- Workflow Value
							(SELECT IV.value
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements WE ON IV.elementID = WE.ID
							INNER JOIN ASRSysWorkflowElements WE2 ON WE.workflowID = WE2.workflowID
							WHERE WE.identifier = EC.WFFormIdentifier
								AND WE2.id = @piElementID
								AND IV.instanceID = @piInstanceID
								AND IV.identifier = EC.WFValueIdentifier)
						ELSE '''' -- Database Value. Handle below to avoid collation conflict.
						END AS [value], 
						EC.valueType, 
						EC.ID,
						EC.calcID,
						isnull(SC.size, 0),
						isnull(SC.decimals, 0)
				FROM ASRSysWorkflowElementColumns EC
				INNER JOIN ASRSysColumns SC ON EC.columnID = SC.columnID
				WHERE EC.elementID = @piElementID
					AND ((SC.dataType <> -3) AND (SC.dataType <> -4));
		
				OPEN columnCursor;
				FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID, @iCalcID, @iSize, @iDecimals;
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @iValueType = 2 -- DBValue - get here to avoid collation conflict
					BEGIN
						SELECT @sValue = dbV.value
						FROM @dbValues dbV
						WHERE dbV.ID = @iSDColumnID;
					END
		
					IF @iValueType = 3 -- Calculated Value
					BEGIN
						EXEC [dbo].[spASRSysWorkflowCalculation]
							@piInstanceID,
							@iCalcID,
							@iResultType OUTPUT,
							@sResult OUTPUT,
							@fResult OUTPUT,
							@dtResult OUTPUT,
							@fltResult OUTPUT, 
							0;
		
						IF @iColumnDataType = 12 SET @sResult = LEFT(@sResult, @iSize); -- Character
						IF @iColumnDataType = 2 -- Numeric
						BEGIN
							IF @fltResult >= power(10, @iSize - @iDecimals) SET @fltResult = 0;
							IF @fltResult <= (-1 * power(10, @iSize - @iDecimals)) SET @fltResult = 0;
						END
		
						SET @sValue = 
							CASE
								WHEN @iResultType = 2 THEN ltrim(rtrim(STR(@fltResult, 8000, @iDecimals)))
								WHEN @iResultType = 3 THEN 
									CASE 
										WHEN @fResult = 1 THEN ''1''
										ELSE ''0''
									END
								WHEN (@iResultType = 4) THEN
									CASE 
										WHEN @dtResult is NULL THEN ''NULL''
										ELSE convert(varchar(100), @dtResult, 101)
									END
								ELSE convert(varchar(MAX), @sResult)
							END;
					END
		
					IF @piDataAction = 0 
					BEGIN
						/* INSERT. */
						SET @sColumnList = @sColumnList
							+ CASE
								WHEN LEN(@sColumnList) > 0 THEN '',''
								ELSE ''''
							END
							+ @sColumnName;
		
						SET @sValueList = @sValueList
							+ CASE
								WHEN LEN(@sValueList) > 0 THEN '',''
								ELSE ''''
							END
							+ CASE
								WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''''''' + replace(isnull(@sValue, ''''), '''''''', '''''''''''') + '''''''' -- 12 = varchar, -1 = working pattern
								WHEN @iColumnDataType = 11 THEN
									CASE 
										WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS null) THEN ''null''
										ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
									END
								WHEN LEN(@sValue) = 0 THEN ''0''
								ELSE isnull(@sValue, 0) -- integer, logic, numeric
							END;
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
								WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''''''' + replace(isnull(@sValue, ''''), '''''''', '''''''''''') + '''''''' -- 12 = varchar, -1 = working pattern
								WHEN @iColumnDataType = 11 THEN
									CASE 
										WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS null) THEN ''null''
										ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
									END
								WHEN LEN(@sValue) = 0 THEN ''0''
								ELSE isnull(@sValue, 0) -- integer, logic, numeric
							END;
					END
		
					DELETE FROM [dbo].[ASRSysWorkflowInstanceValues]
					WHERE instanceID = @piInstanceID
						AND elementID = @piElementID
						AND columnID = @iColumnID;
		
					INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
						(instanceID, elementID, identifier, columnID, value, emailID)
						VALUES (@piInstanceID, @piElementID, '''', @iColumnID, @sValue, 0);
		
					FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID, @iCalcID, @iSize, @iDecimals;
				END
		
				CLOSE columnCursor;
				DEALLOCATE columnCursor;
		
				IF @piDataAction = 0 
				BEGIN
					/* INSERT. */
					IF @iDataRecord <> 3 -- 3 = Unidentified record
					BEGIN
						SET @sColumnList = @sColumnList
							+ CASE
								WHEN LEN(@sColumnList) > 0 THEN '',''
								ELSE ''''
							END
							+ @sIDColumnName;
			
						SET @sValueList = @sValueList
							+ CASE
								WHEN LEN(@sValueList) > 0 THEN '',''
								ELSE ''''
							END
							+ convert(varchar(255), @piRecordID);
		
						IF @piDataAction = 0 -- Insert
							AND (@iSecondaryDataRecord = 0 -- 0 = Initiator''s record
								OR @iSecondaryDataRecord = 1 -- 1 = Previous record selector''s record
								OR @iSecondaryDataRecord = 4) -- 4 = Triggered record
						BEGIN
							SET @sColumnList = @sColumnList
								+ CASE
									WHEN LEN(@sColumnList) > 0 THEN '',''
									ELSE ''''
								END
								+ @sSecondaryIDColumnName;
						
							SET @sValueList = @sValueList
								+ CASE
									WHEN LEN(@sValueList) > 0 THEN '',''
									ELSE ''''
								END
								+ convert(varchar(255), @iSecondaryRecordID);
						END
					END
				END
		
				IF LEN(@sColumnList) > 0
				BEGIN
					IF @piDataAction = 0 
					BEGIN
						/* INSERT. */
						SET @psSQL = ''INSERT INTO '' + @psTableName
							+ '' ('' + @sColumnList + '')''
							+ '' VALUES('' + @sValueList + '')'';
					END
					ELSE
					BEGIN
						/* UPDATE. */
						SET @psSQL = ''UPDATE '' + @psTableName
							+ '' SET '' + @sColumnList
							+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(255), @piRecordID);
					END
				END
			END
		
			IF @piDataAction = 2
			BEGIN
				/* DELETE. */
				SET @psSQL = ''DELETE FROM '' + @psTableName
					+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(255), @piRecordID);
			END	
		
			IF (@piDataAction = 0) -- Insert
			BEGIN
				SET @iParent1TableID = isnull(@iDataRecordTableID, 0);
				SET @iParent1RecordID = isnull(@piRecordID, 0);
				SET @iParent2TableID = isnull(@iSecondaryDataRecordTableID, 0);
				SET @iParent2RecordID = isnull(@iSecondaryRecordID, 0);
			END
			ELSE
			BEGIN	-- Update or Delete
				exec [dbo].[spASRGetParentDetails]
					@piDataTableID,
					@piRecordID,
					@iParent1TableID	OUTPUT,
					@iParent1RecordID	OUTPUT,
					@iParent2TableID	OUTPUT,
					@iParent2RecordID	OUTPUT;
			END
		
			UPDATE ASRSysWorkflowInstanceValues
			SET ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID, 
				ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
				ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID, 
				ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
			WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceValues.elementID = @piElementID
				AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
				AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
		
			IF (@piDataAction = 2) -- Delete
			BEGIN
				DECLARE curColumns CURSOR LOCAL FAST_FORWARD FOR 
				SELECT columnID
				FROM [dbo].[udfASRWorkflowColumnsUsed] (@iWorkflowID, @piElementID, 0);
		
				OPEN curColumns;
		
				FETCH NEXT FROM curColumns INTO @iDBColumnID;
				WHILE (@@fetch_status = 0)
				BEGIN
					DELETE FROM ASRSysWorkflowInstanceValues
					WHERE instanceID = @piInstanceID
						AND elementID = @piElementID
						AND columnID = @iDBColumnID;
		
					SELECT @sDBTableName = tbl.tableName,
						@iRequiredTableID = tbl.tableID, 
						@sDBColumnName = col.columnName,
						@iDataType = col.dataType
					FROM ASRSysColumns col
					INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
					WHERE col.columnID = @iDBColumnID;
		
					SET @sSQL = ''SELECT @sDBValue = ''
						+ CASE
							WHEN @iDataType = 12 THEN ''''
							WHEN @iDataType = 11 THEN ''convert(varchar(MAX),''
							ELSE ''convert(varchar(MAX),''
						END
						+ @sDBTableName + ''.'' + @sDBColumnName
						+ CASE
							WHEN @iDataType = 12 THEN ''''
							WHEN @iDataType = 11 THEN '', 101)''
							ELSE '')''
						END
						+ '' FROM '' + @sDBTableName 
						+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(255), @piRecordID);
		
					SET @sParam = N''@sDBValue varchar(MAX) OUTPUT'';
					EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;
		
					INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
						(instanceID, elementID, identifier, columnID, value, emailID)
						VALUES (@piInstanceID, @piElementID, '''', @iDBColumnID, @sDBValue, 0);
							
					FETCH NEXT FROM curColumns INTO @iDBColumnID;
				END
				CLOSE curColumns;
				DEALLOCATE curColumns;
		
				DECLARE curEmails CURSOR LOCAL FAST_FORWARD FOR 
				SELECT emailID,
					type,
					colExprID
				FROM [dbo].[udfASRWorkflowEmailsUsed] (@iWorkflowID, @piElementID, 0);
		
				OPEN curEmails;
		
				FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID;
				WHILE (@@fetch_status = 0)
				BEGIN
					DELETE FROM [dbo].[ASRSysWorkflowInstanceValues]
					WHERE instanceID = @piInstanceID
						AND elementID = @piElementID
						AND emailID = @iEmailID;
		
					IF @iType = 1 -- Column
					BEGIN
						SELECT @sDBTableName = tbl.tableName,
							@iRequiredTableID = tbl.tableID, 
							@sDBColumnName = col.columnName,
							@iDataType = col.dataType
						FROM [dbo].[ASRSysColumns] col
						INNER JOIN [dbo].[ASRSysTables] tbl ON col.tableID = tbl.tableID
						WHERE col.columnID = @iDBColumnID;
		
						SET @sSQL = ''SELECT @sDBValue = ''
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN ''convert(varchar(MAX),''
								ELSE ''convert(varchar(MAX),''
							END
							+ @sDBTableName + ''.'' + @sDBColumnName
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN '', 101)''
								ELSE '')''
							END
							+ '' FROM '' + @sDBTableName 
							+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(255), @piRecordID);
		
						SET @sParam = N''@sDBValue varchar(MAX) OUTPUT'';
						EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;
					END
					ELSE
					BEGIN
						EXEC [dbo].[spASRSysEmailAddr]
							@sDBValue OUTPUT,
							@iEmailID,
							@piRecordID;
					END
		
					INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
						(instanceID, elementID, identifier, columnID, value, emailID)
						VALUES (@piInstanceID, @piElementID, '''', 0, @sDBValue, @iEmailID);
							
					FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID;
				END
				CLOSE curEmails;
				DEALLOCATE curEmails;
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetWindowsUsers
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWindowsUsers]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWindowsUsers];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetWindowsUsers]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetWindowsUsers]
		(
			@DomainName varchar(200),
			@UserString varchar(MAX) OUTPUT
		)
		AS
		BEGIN
			SELECT @UserString = dbo.udfASRNetGetUsers(@DomainName);
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetWorkflowDelegates
	----------------------------------------------------------------------

	-- JIRA367 - Strange issues with collation sequences. (Procedure regened during SS save)
	IF NOT OBJECT_ID('udfASRGetWorkflowDelegatedRecords', 'TF') IS NULL	
		DROP FUNCTION [dbo].[udfASRGetWorkflowDelegatedRecords]

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowDelegates]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowDelegates];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowDelegates]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowDelegates] 
		(
			@psTo			varchar(MAX),
			@piStepID		integer,
			@results		cursor varying output
		)
		AS
		BEGIN
			DECLARE
				@iDelegateEmailID	integer,
				@sTemp				varchar(MAX),
				@iDelegateRecordID	integer,
				@sDelegateTo		varchar(MAX),
				@iCount				integer,
				@sSQL				nvarchar(MAX),
				@iInstanceID		integer;
		
			IF len(ltrim(rtrim(@psTo))) = 0 RETURN;
		
		    DECLARE @recipients TABLE (
		        recordID		integer,
				emailAddress	varchar(MAX),
				delegated		bit,
				delegatedTo		varchar(MAX),
				processed		tinyint default 0,
				isDelegate		bit);
				
			-- Get the delegate email address definition. 
			SET @iDelegateEmailID = 0;
			SELECT @sTemp = ISNULL(parameterValue, '''')
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_DelegateEmail''
			SET @iDelegateEmailID = convert(integer, @sTemp);
				
			IF @iDelegateEmailID = 0
			BEGIN
				INSERT INTO @recipients (
					recordID,
					emailAddress,
					delegated,
					delegatedTo,
					processed,
					isDelegate)
				VALUES (
					0, -- Personnel Record ID
					@psTo, -- Email Address(es)
					0, -- Delegated
					'''', -- Delegate Email Address
					2, -- Processed
					0); -- Is Delegate
			END
			ELSE
			BEGIN
				INSERT INTO @recipients 
				SELECT         
					RECnew.ID,
					RECnew.emailAddress,
					RECnew.delegated,
					RECnew.delegatedTo,
					0,
					0
				FROM [dbo].[udfASRGetWorkflowDelegatedRecords](@psTo) RECnew
				WHERE len(ltrim(rtrim(RECnew.emailAddress))) > 0
					AND RECnew.emailAddress NOT IN (SELECT RECold.emailAddress 
						FROM @recipients RECold
						WHERE RECold.recordID = 0 OR RECold.recordID = RECnew.ID);
		
				SELECT @iCount = COUNT(*)
				FROM @recipients
				WHERE processed = 0;
		
				WHILE @iCount > 0
				BEGIN
					-- Mark the new rows as ''being processed''.
					UPDATE @recipients
					SET processed = 1
					WHERE processed = 0;
		
					DECLARE delegatesCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT recordID
					FROM @recipients
					WHERE recordID > 0
						AND processed = 1
						AND delegated = 1;
		
					OPEN delegatesCursor;
					FETCH NEXT FROM delegatesCursor INTO @iDelegateRecordID;
					WHILE (@@fetch_status = 0)
					BEGIN
						SET @sDelegateTo = '''';
						SET @sSQL = ''spASRSysEmailAddr'';
			
						IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
						BEGIN
							-- Get the delegate''s email address
							EXEC @sSQL @sDelegateTo OUTPUT, @iDelegateEmailID, @iDelegateRecordID;
							IF @sDelegateTo IS null SET @sDelegateTo = '''';
						END
		
						IF len(@sDelegateTo) > 0 
						BEGIN
							UPDATE @recipients 
							SET delegatedTo = @sDelegateTo
							WHERE recordID = @iDelegateRecordID;
		
							INSERT INTO @recipients 
							SELECT         
								RECnew.ID,
								RECnew.emailAddress,
								RECnew.delegated,
								RECnew.delegatedTo,
								0,
								1
							FROM [dbo].[udfASRGetWorkflowDelegatedRecords](@sDelegateTo) RECnew
							WHERE len(ltrim(rtrim(RECnew.emailAddress))) > 0
								AND RECnew.emailAddress NOT IN (SELECT RECold.emailAddress 
									FROM @recipients RECold
									WHERE RECold.recordID = 0 OR RECold.recordID = RECnew.ID);
						END
						ELSE
						BEGIN
							UPDATE @recipients 
							SET delegated = 0
							WHERE recordID = @iDelegateRecordID;
						END
		
						FETCH NEXT FROM delegatesCursor INTO @iDelegateRecordID;
					END
					CLOSE delegatesCursor;
					DEALLOCATE delegatesCursor;
		
					-- Mark the processed rows as ''been processed''.
					UPDATE @recipients
					SET processed = 2
					WHERE processed = 1;
		
					SELECT @iCount = COUNT(*)
					FROM @recipients
					WHERE processed = 0;
				END
			END
		
			-- Return the cursor of succeeding elements. 
			SET @results = CURSOR FORWARD_ONLY STATIC FOR
		        SELECT DISTINCT 
					emailAddress,
					delegated,
					delegatedTo,
					isDelegate
		        FROM @recipients;
		
			OPEN @results;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetWorkflowEmailMessage
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowEmailMessage]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowEmailMessage];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowEmailMessage]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowEmailMessage]
		(
			@piInstanceID					integer,
			@piElementID					integer,
			@psMessage						varchar(MAX)	OUTPUT, 
			@psMessage_HypertextLinks		varchar(MAX)	OUTPUT, 
			@psHypertextLinkedSteps			varchar(MAX)	OUTPUT, 
			@pfOK							bit	OUTPUT,
			@psTo							varchar(MAX)
		)
		AS
		BEGIN
			DECLARE 
				@iInitiatorID		integer,
				@sCaption			varchar(MAX),
				@iItemType			integer,
				@iDBColumnID		integer,
				@iDBRecord			integer,
				@sWFFormIdentifier	varchar(MAX),
				@sWFValueIdentifier	varchar(MAX),
				@sValue				varchar(MAX),
				@sTemp				varchar(MAX),
				@sTableName			sysname,
				@sColumnName		sysname,
				@iRecordID			integer,
				@sSQL				nvarchar(MAX),
				@sSQLParam			nvarchar(MAX),
				@iCount				integer,
				@iElementID			integer,
				@superCursor		cursor,
				@iTemp				integer,
				@hResult 			integer,
				@objectToken 		integer,
				@sQueryString		varchar(MAX),
				@sURL				varchar(MAX), 
				@sEmailFormat		varchar(MAX),
				@iEmailFormat		integer,
				@iSourceItemType	integer,
				@dtTempDate			datetime, 
				@sParam1			varchar(MAX),
				@sDBName			sysname,
				@sRecSelWebFormIdentifier	varchar(MAX),
				@sRecSelIdentifier	varchar(MAX),
				@iElementType		integer,
				@iWorkflowID		integer, 
				@fValidRecordID		bit,
				@iBaseTableID		integer,
				@iBaseRecordID		integer,
				@iRequiredTableID	integer,
				@iRequiredRecordID	integer,
				@iParent1TableID	integer,
				@iParent1RecordID	integer,
				@iParent2TableID	integer,
				@iParent2RecordID	integer,
				@iInitParent1TableID	integer,
				@iInitParent1RecordID	integer,
				@iInitParent2TableID	integer,
				@iInitParent2RecordID	integer,
				@fDeletedValue		bit,
				@iTempElementID		integer,
				@iColumnID			integer,
				@iResultType		integer,
				@sResult			varchar(MAX),
				@fResult			bit,
				@dtResult			datetime,
				@fltResult			float,
				@iCalcID			integer,
				@iPersonnelTableID	integer,
				@iSQLVersion		integer;
						
			SET @pfOK = 1;
			SET @psMessage = '''';
			SET @psMessage_HypertextLinks = '''';
			SET @psHypertextLinkedSteps = '''';
			SELECT @iSQLVersion = dbo.udfASRSQLVersion();
		
			SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_PERSONNEL''
				AND parameterKey = ''Param_TablePersonnel'';
		
			IF @iPersonnelTableID = 0
			BEGIN
				SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_TablePersonnel'';
			END
		
			exec [dbo].[spASRGetSetting]
				''email'',
				''date format'',
				''103'',
				0,
				@sEmailFormat OUTPUT;
		
			SET @iEmailFormat = convert(integer, @sEmailFormat);
			
			SELECT @sURL = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_URL'';
		
			IF upper(right(@sURL, 5)) <> ''.ASPX''
				AND right(@sURL, 1) <> ''/''
				AND len(@sURL) > 0
			BEGIN
				SET @sURL = @sURL + ''/'';
			END
		
			SELECT @sParam1 = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''		
				AND parameterKey = ''Param_Web1'';
			
			SET @sDBName = db_name()
		
			SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
				@iWorkflowID = ASRSysWorkflowInstances.workflowID,
				@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
				@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
				@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
				@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
			FROM ASRSysWorkflowInstances
			WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
		
			DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT EI.caption,
				EI.itemType,
				EI.dbColumnID,
				EI.dbRecord,
				EI.wfFormIdentifier,
				EI.wfValueIdentifier, 
				EI.recSelWebFormIdentifier,
				EI.recSelIdentifier, 
				EI.calcID
			FROM ASRSysWorkflowElementItems EI
			WHERE EI.elementID = @piElementID
			ORDER BY EI.ID;
		
			OPEN itemCursor;
			FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier, @iCalcID;
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @sValue = '''';
		
				IF @iItemType = 1
				BEGIN
					SET @fDeletedValue = 0;
		
					/* Database value. */
					SELECT @sTableName = ASRSysTables.tableName, 
						@iRequiredTableID = ASRSysTables.tableID, 
						@sColumnName = ASRSysColumns.columnName, 
						@iSourceItemType = ASRSysColumns.dataType
					FROM ASRSysColumns
					INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
					WHERE ASRSysColumns.columnID = @iDBColumnID;
		
					IF @iDBRecord = 0
					BEGIN
						-- Initiator''s record
						SET @iRecordID = @iInitiatorID;
						SET @iParent1TableID = @iInitParent1TableID;
						SET @iParent1RecordID = @iInitParent1RecordID;
						SET @iParent2TableID = @iInitParent2TableID;
						SET @iParent2RecordID = @iInitParent2RecordID;
						SET @iBaseTableID = @iPersonnelTableID;
					END			
		
					IF @iDBRecord = 4
					BEGIN
						-- Trigger record
						SET @iRecordID = @iInitiatorID;
						SET @iParent1TableID = @iInitParent1TableID;
						SET @iParent1RecordID = @iInitParent1RecordID;
						SET @iParent2TableID = @iInitParent2TableID;
						SET @iParent2RecordID = @iInitParent2RecordID;
		
						SELECT @iBaseTableID = isnull(WF.baseTable, 0)
						FROM ASRSysWorkflows WF
						INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
							AND WFI.ID = @piInstanceID;
					END
		
					IF @iDBRecord = 1
					BEGIN
						-- Previously identified record.
						SELECT @iElementType = ASRSysWorkflowElements.type, 
							@iTempElementID = ASRSysWorkflowElements.ID
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)));
		
						IF @iElementType = 2
						BEGIN
							 -- WebForm
							SELECT @sTemp = ISNULL(IV.value, ''0''),
								@iBaseTableID = EI.tableID,
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
							INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
							WHERE IV.instanceID = @piInstanceID
								AND IV.identifier = @sRecSelIdentifier
								AND Es.identifier = @sRecSelWebFormIdentifier
								AND Es.workflowID = @iWorkflowID
								AND IV.elementID = Es.ID;
						END
						ELSE
						BEGIN
							-- StoredData
							SELECT @sTemp = ISNULL(IV.value, ''0''),
								@iBaseTableID = isnull(Es.dataTableID, 0),
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
								AND IV.identifier = Es.identifier
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sRecSelWebFormIdentifier
							WHERE IV.instanceID = @piInstanceID;
						END
		
						SET @iRecordID = 
							CASE
								WHEN isnumeric(@sTemp) = 1 THEN convert(integer, @sTemp)
								ELSE 0
							END;
					END		
		
					SET @iBaseRecordID = @iRecordID;
		
					IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
					BEGIN
						SET @fValidRecordID = 0;
		
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
							@iBaseTableID,
							@iBaseRecordID,
							@iParent1TableID,
							@iParent1RecordID,
							@iParent2TableID,
							@iParent2RecordID,
							@iRequiredTableID,
							@iRequiredRecordID	OUTPUT;
		
						SET @iRecordID = @iRequiredRecordID
		
						IF @iRecordID > 0 
						BEGIN
							EXEC [dbo].[spASRWorkflowValidTableRecord]
								@iRequiredTableID,
								@iRecordID,
								@fValidRecordID	OUTPUT;
						END
		
						IF @fValidRecordID = 0
						BEGIN
							IF @iDBRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
							BEGIN
								SELECT @iCount = COUNT(*)
								FROM ASRSysWorkflowQueueColumns QC
								INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
								WHERE WFQ.instanceID = @piInstanceID
									AND QC.columnID = @iDBColumnID;
		
								IF @iCount = 1
								BEGIN
									SELECT @sValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
									FROM ASRSysWorkflowQueueColumns QC
									INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
									WHERE WFQ.instanceID = @piInstanceID
										AND QC.columnID = @iDBColumnID;
		
									SET @fValidRecordID = 1;
									SET @fDeletedValue = 1;
								END
							END
							ELSE
							BEGIN
								IF @iDBRecord = 1
								BEGIN
									SELECT @iCount = COUNT(*)
									FROM ASRSysWorkflowInstanceValues IV
									WHERE IV.instanceID = @piInstanceID
										AND IV.columnID = @iDBColumnID
										AND IV.elementID = @iTempElementID;
		
									IF @iCount = 1
									BEGIN
										SELECT @sValue = rtrim(ltrim(isnull(IV.value , '''')))
										FROM ASRSysWorkflowInstanceValues IV
										WHERE IV.instanceID = @piInstanceID
											AND IV.columnID = @iDBColumnID
											AND IV.elementID = @iTempElementID;
		
										SET @fValidRecordID = 1;
										SET @fDeletedValue = 1;
									END
								END
							END
						END
		
						IF @fValidRecordID  = 0
						BEGIN
							SET @psMessage = '''';
							SET @pfOK = 0;
		
							RETURN;
						END
					END
		
					IF @fDeletedValue = 0
					BEGIN
						SET @sSQL = ''SELECT @sValue = '' + @sTableName + ''.'' + @sColumnName +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(255), @iRecordID);
						SET @sSQLParam = N''@sValue varchar(MAX) OUTPUT'';
						EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT;
					END					
					IF @sValue IS null SET @sValue = '''';
		
					/* Format dates */
					IF @iSourceItemType = 11
					BEGIN
						IF len(@sValue) = 0
						BEGIN
							SET @sValue = ''<undefined>'';
						END
						ELSE
						BEGIN
							SET @dtTempDate = convert(datetime, @sValue);
							SET @sValue = convert(varchar(MAX), @dtTempDate, @iEmailFormat);
						END
					END
		
					/* Format logics */
					IF @iSourceItemType = -7
					BEGIN
						IF @sValue = 0 
						BEGIN
							SET @sValue = ''False'';
						END
						ELSE
						BEGIN
							SET @sValue = ''True'';
						END
					END	
		
					SET @psMessage = @psMessage
						+ @sValue;
				END
				
				IF @iItemType = 2
				BEGIN
					/* Label value. */
					SET @psMessage = @psMessage
						+ @sCaption;
				END
		
				IF @iItemType = 4
				BEGIN
					/* Workflow value. */
					SELECT @sValue = ASRSysWorkflowInstanceValues.value, 
						@iSourceItemType = ASRSysWorkflowElementItems.itemType,
						@iColumnID = ASRSysWorkflowElementItems.lookupColumnID
					FROM ASRSysWorkflowInstanceValues
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
					INNER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.elementID
					WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
						AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
						AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowElementItems.identifier = @sWFValueIdentifier;
		
					IF @sValue IS null SET @sValue = '''';
		
					IF @iSourceItemType = 14 -- Lookup, need to get the column data type
					BEGIN
						SELECT @iSourceItemType = 
							CASE
								WHEN ASRSysColumns.dataType = -7 THEN 6 -- Logic
								WHEN ASRSysColumns.dataType = 2 THEN 5 -- Numeric
								WHEN ASRSysColumns.dataType = 4 THEN 5 -- Integer
								WHEN ASRSysColumns.dataType = 11 THEN 7 -- Date
								ELSE 3
							END
						FROM ASRSysColumns
						WHERE ASRSysColumns.columnID = @iColumnID;
					END
							
					/* Format dates */
					IF @iSourceItemType = 7
					BEGIN
						IF len(@sValue) = 0 OR @sValue = ''null''
						BEGIN
							SET @sValue = ''<undefined>'';
						END
						ELSE
						BEGIN
							SET @dtTempDate = convert(datetime, @sValue);
							SET @sValue = convert(varchar(MAX), @dtTempDate, @iEmailFormat);
						END
					END
		
					/* Format logics */
					IF @iSourceItemType = 6
					BEGIN
						IF @sValue = 0 
						BEGIN
							SET @sValue = ''False'';
						END
						ELSE
						BEGIN
							SET @sValue = ''True'';
						END
					END			
		
					SET @psMessage = @psMessage
						+ @sValue;
				END
		
				IF @iItemType = 12
				BEGIN
					/* Formatting option. */
					/* NB. The empty string that precede the char codes ARE required. */
					SET @psMessage = @psMessage +
						CASE
							WHEN @sCaption = ''L'' THEN '''' + char(13) + char(10) + ''--------------------------------------------------'' + char(13) + char(10)
							WHEN @sCaption = ''N'' THEN '''' + char(13) + char(10)
							WHEN @sCaption = ''T'' THEN '''' + char(9)
							ELSE ''''
						END;
				END
		
				IF @iItemType = 16
				BEGIN
					/* Calculation. */
					EXEC [dbo].[spASRSysWorkflowCalculation]
						@piInstanceID,
						@iCalcID,
						@iResultType OUTPUT,
						@sResult OUTPUT,
						@fResult OUTPUT,
						@dtResult OUTPUT,
						@fltResult OUTPUT, 
						0;
		
					SET @psMessage = @psMessage +
						@sResult;
				END
		
				FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier, @iCalcID;
			END
			CLOSE itemCursor;
			DEALLOCATE itemCursor;
		
			/* Append the link to the webform that follows this element (ignore connectors) if there are any. */
			CREATE TABLE #succeedingElements (elementID integer);
		
			EXEC [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]  
				@piInstanceID, 
				@piElementID, 
				@superCursor OUTPUT,
				@psTo;
		
			FETCH NEXT FROM @superCursor INTO @iTemp;
			WHILE (@@fetch_status = 0)
			BEGIN
				INSERT INTO #succeedingElements (elementID) VALUES (@iTemp);
				
				FETCH NEXT FROM @superCursor INTO @iTemp;
			END
			CLOSE @superCursor;
			DEALLOCATE @superCursor;
		
			SELECT @iCount = COUNT(*)
			FROM #succeedingElements SE
			INNER JOIN ASRSysWorkflowElements WE ON SE.elementID = WE.id
			WHERE WE.type = 2; -- 2 = Web Form element
		
			IF @iCount > 0 
			BEGIN
				SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + char(13) + char(10) + char(13) + char(10)
					+ ''Click on the following link''
					+ CASE
						WHEN @iCount = 1 THEN ''''
						ELSE ''s''
					END
					+ '' to action:''
					+ char(13) + char(10);
		
				DECLARE elementCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT SE.elementID, ISNULL(WE.caption, '''')
				FROM #succeedingElements SE
				INNER JOIN ASRSysWorkflowElements WE ON SE.elementID = WE.ID
				WHERE WE.type = 2; -- 2 = Web Form element
			
				OPEN elementCursor;
				FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption;
				WHILE (@@fetch_status = 0)
				BEGIN
		
					SELECT @sQueryString = dbo.[udfASRNetGetWorkflowQueryString]
						(@piInstanceID, @iElementID, @sParam1, @@servername, @sDBName);
								
					IF LEN(@sQueryString) = 0 
					BEGIN
						SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + char(13) + char(10) +
							@sCaption + '' - Error constructing the query string. Please contact your system administrator.'';
					END
					ELSE
					BEGIN
						SET @psHypertextLinkedSteps = @psHypertextLinkedSteps
							+ CASE
								WHEN len(@psHypertextLinkedSteps) = 0 THEN char(9)
								ELSE ''''
							END 
							+ convert(varchar(MAX), @iElementID)
							+ char(9);
		
						SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + char(13) + char(10) +
							@sCaption + '' - '' + char(13) + char(10) + 
							''<'' + @sURL + ''?'' + @sQueryString + ''>'';
					END
					
					FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption;
				END
		
				CLOSE elementCursor;
				DEALLOCATE elementCursor;
		
				SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + char(13) + char(10) + char(13) + char(10)
					+ ''Please make sure that the link''
					+ CASE
						WHEN @iCount = 1 THEN '' has''
						ELSE ''s have''
					END
					+ '' not been cut off by your display.'' + char(13) + char(10)
					+ ''If ''
					+ CASE
						WHEN @iCount = 1 THEN ''it has''
						ELSE ''they have''
					END
					+ '' been cut off you will need to copy and paste ''
					+ CASE
						WHEN @iCount = 1 THEN ''it''
						ELSE ''them''
					END
					+ '' into your browser.'';
			END
		
			DROP TABLE #succeedingElements;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetWorkflowFileUploadDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowFileUploadDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowFileUploadDetails];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowFileUploadDetails]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowFileUploadDetails]
		(
			@piElementItemID	integer,
			@piInstanceID		integer,
			@piSize				integer			OUTPUT,
			@psFileName			varchar(MAX)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iElementID		integer,
				@sIdentifier	varchar(MAX) 
		
			SELECT 			
				@piSize = ISNULL(ASRSysWorkflowElementItems.InputSize, 0),
				@iElementID = elementID,
				@sIdentifier = identifier
			FROM ASRSysWorkflowElementItems
			WHERE ASRSysWorkflowElementItems.ID = @piElementItemID;
		
			SELECT @psFileName = [TempFileUpload_Filename]
			FROM ASRSysWorkflowInstanceValues
			WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceValues.elementID = @iElementID
				AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
		
			SELECT ASRSysWorkflowElementItemValues.value
			FROM ASRSysWorkflowElementItemValues
			WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetWorkflowFormItems
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowFormItems]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowFormItems];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowFormItems]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowFormItems]
		(
			@piInstanceID				integer,
			@piElementID				integer,
			@psErrorMessage				varchar(MAX)	OUTPUT,
			@piBackColour				integer			OUTPUT,
			@piBackImage				integer			OUTPUT,
			@piBackImageLocation		integer			OUTPUT,
			@piWidth					integer			OUTPUT,
			@piHeight					integer			OUTPUT,
			@piCompletionMessageType	integer			OUTPUT,
			@psCompletionMessage		varchar(200)	OUTPUT,
			@piSavedForLaterMessageType	integer			OUTPUT,
			@psSavedForLaterMessage		varchar(200)	OUTPUT,
			@piFollowOnFormsMessageType	integer			OUTPUT,
			@psFollowOnFormsMessage		varchar(200)	OUTPUT
		)
		AS
		BEGIN
			DECLARE 
				@iID				integer,
				@iItemType			integer,
				@iDefaultValueType	integer,
				@iDBColumnID		integer,
				@iDBColumnDataType	integer,
				@iDBRecord			integer,
				@sWFFormIdentifier	varchar(MAX),
				@sWFValueIdentifier	varchar(MAX),
				@sValue				varchar(MAX),
				@sSQL				nvarchar(MAX),
				@sSQLParam			nvarchar(500),
				@sTableName			sysname,
				@sColumnName		sysname,
				@iInitiatorID		integer,
				@iRecordID			integer,
				@iStatus			integer,
				@iCount				integer,
				@iWorkflowID		integer,
				@iElementType		integer, 
				@iType				integer,
				@fValidRecordID		bit,
				@iBaseTableID		integer,
				@iBaseRecordID		integer,
				@iRequiredTableID	integer,
				@iRequiredRecordID	integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@iInitParent1TableID	integer,
				@iInitParent1RecordID	integer,
				@iInitParent2TableID	integer,
				@iInitParent2RecordID	integer,
				@fDeletedValue			bit,
				@iTempElementID			integer,
				@iColumnID				integer,
				@iResultType			integer,
				@sResult				varchar(MAX),
				@fResult				bit,
				@dtResult				datetime,
				@fltResult				float,
				@iCalcID				integer,
				@iSize					integer,
				@iDecimals				integer,
				@iPersonnelTableID		integer,
				@sIdentifier			varchar(MAX);
		
			DECLARE @itemValues table(ID integer, value varchar(MAX), type integer)	
					
			-- Check the given instance still exists.
			SELECT @iCount = COUNT(*)
			FROM ASRSysWorkflowInstances
			WHERE ASRSysWorkflowInstances.ID = @piInstanceID
		
			IF @iCount = 0
			BEGIN
				SET @psErrorMessage = ''This workflow step is invalid. The workflow process may have been completed.''
				RETURN
			END
		
			-- Check if the step has already been completed!
			SELECT @iStatus = ASRSysWorkflowInstanceSteps.status
			FROM ASRSysWorkflowInstanceSteps
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
		
			IF @iStatus = 3
			BEGIN
				SET @psErrorMessage = ''This workflow step has already been completed.''
				RETURN
			END
		
			IF @iStatus = 6
			BEGIN
				SET @psErrorMessage = ''This workflow step has timed out.''
				RETURN
			END
		
			IF @iStatus = 0
			BEGIN
				SET @psErrorMessage = ''This workflow step is invalid. It may no longer be required due to the results of other workflow steps.''
				RETURN
			END
		
			SET @psErrorMessage = ''''
		
			SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_PERSONNEL''
				AND parameterKey = ''Param_TablePersonnel''
		
			IF @iPersonnelTableID = 0
			BEGIN
				SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_TablePersonnel''
			END
						
			SELECT 			
				@piBackColour = isnull(webFormBGColor, 16777166),
				@piBackImage = isnull(webFormBGImageID, 0),
				@piBackImageLocation = isnull(webFormBGImageLocation, 0),
				@piWidth = isnull(webFormWidth, -1),
				@piHeight = isnull(webFormHeight, -1),
				@iWorkflowID = workflowID,
				@piCompletionMessageType = CompletionMessageType,
				@psCompletionMessage = CompletionMessage,
				@piSavedForLaterMessageType = SavedForLaterMessageType,
				@psSavedForLaterMessage = SavedForLaterMessage,
				@piFollowOnFormsMessageType = FollowOnFormsMessageType,
				@psFollowOnFormsMessage = FollowOnFormsMessage
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.ID = @piElementID
		
			SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
				@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
				@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
				@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
				@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
			FROM ASRSysWorkflowInstances
			WHERE ASRSysWorkflowInstances.ID = @piInstanceID
		
			DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysWorkflowElementItems.ID,
				ASRSysWorkflowElementItems.itemType,
				ASRSysWorkflowElementItems.dbColumnID,
				ASRSysWorkflowElementItems.dbRecord,
				ASRSysWorkflowElementItems.wfFormIdentifier,
				ASRSysWorkflowElementItems.wfValueIdentifier,
				ASRSysWorkflowElementItems.calcID,
				ASRSysWorkflowElementItems.identifier,
				isnull(ASRSysWorkflowElementItems.defaultValueType, 0) AS [defaultValueType],
				isnull(ASRSysWorkflowElementItems.inputSize, 0),
				isnull(ASRSysWorkflowElementItems.inputDecimals, 0)
			FROM ASRSysWorkflowElementItems
			WHERE ASRSysWorkflowElementItems.elementID = @piElementID
				AND (ASRSysWorkflowElementItems.itemType = 1 
					OR (ASRSysWorkflowElementItems.itemType = 2 AND ASRSysWorkflowElementItems.captionType = 3)
					OR ASRSysWorkflowElementItems.itemType = 3
					OR ASRSysWorkflowElementItems.itemType = 5
					OR ASRSysWorkflowElementItems.itemType = 6
					OR ASRSysWorkflowElementItems.itemType = 7
					OR ASRSysWorkflowElementItems.itemType = 11
					OR ASRSysWorkflowElementItems.itemType = 17
					OR ASRSysWorkflowElementItems.itemType = 19
					OR ASRSysWorkflowElementItems.itemType = 20
					OR ASRSysWorkflowElementItems.itemType = 4)
		
			OPEN itemCursor
			FETCH NEXT FROM itemCursor INTO 
				@iID, 
				@iItemType, 
				@iDBColumnID, 
				@iDBRecord, 
				@sWFFormIdentifier, 
				@sWFValueIdentifier, 
				@iCalcID, 
				@sIdentifier, 
				@iDefaultValueType,
				@iSize,
				@iDecimals
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @sValue = ''''
		
				IF @iItemType = 1
				BEGIN
					SET @fDeletedValue = 0
		
					-- Database value. 
					SELECT @sTableName = ASRSysTables.tableName, 
						@iRequiredTableID = ASRSysTables.tableID, 
						@sColumnName = ASRSysColumns.columnName,
						@iDBColumnDataType = ASRSysColumns.dataType
					FROM ASRSysColumns
					INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
					WHERE ASRSysColumns.columnID = @iDBColumnID
		
					SET @iType = @iDBColumnDataType
		
					IF @iDBRecord = 0
					BEGIN
						-- Initiator''s record
						SET @iRecordID = @iInitiatorID
						SET @iParent1TableID = @iInitParent1TableID
						SET @iParent1RecordID = @iInitParent1RecordID
						SET @iParent2TableID = @iInitParent2TableID
						SET @iParent2RecordID = @iInitParent2RecordID
						SET @iBaseTableID = @iPersonnelTableID
					END			
		
					IF @iDBRecord = 4
					BEGIN
						-- Trigger record
						SET @iRecordID = @iInitiatorID
						SET @iParent1TableID = @iInitParent1TableID
						SET @iParent1RecordID = @iInitParent1RecordID
						SET @iParent2TableID = @iInitParent2TableID
						SET @iParent2RecordID = @iInitParent2RecordID
		
						SELECT @iBaseTableID = isnull(WF.baseTable, 0)
						FROM ASRSysWorkflows WF
						INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
							AND WFI.ID = @piInstanceID
					END
		
					IF @iDBRecord = 1
					BEGIN
						-- Identified record.
						SELECT @iElementType = ASRSysWorkflowElements.type, 
							@iTempElementID = ASRSysWorkflowElements.ID
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sWFFormIdentifier)))
							
						IF @iElementType = 2
						BEGIN
							 -- WebForm
							SELECT @sValue = ISNULL(IV.value, ''0''),
								@iBaseTableID = EI.tableID,
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
							INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
							WHERE IV.instanceID = @piInstanceID
								AND IV.identifier = @sWFValueIdentifier
								AND Es.identifier = @sWFFormIdentifier
								AND Es.workflowID = @iWorkflowID
								AND IV.elementID = Es.ID
						END
						ELSE
						BEGIN
							-- StoredData
							SELECT @sValue = ISNULL(IV.value, ''0''),
								@iBaseTableID = isnull(Es.dataTableID, 0),
								@iParent1TableID = IV.parent1TableID,
								@iParent1RecordID = IV.parent1RecordID,
								@iParent2TableID = IV.parent2TableID,
								@iParent2RecordID = IV.parent2RecordID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
								AND IV.identifier = Es.identifier
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sWFFormIdentifier
							WHERE IV.instanceID = @piInstanceID
						END
		
						SET @iRecordID = 
							CASE
								WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
								ELSE 0
							END
					END	
					
					SET @iBaseRecordID = @iRecordID
		
					IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
					BEGIN
						SET @fValidRecordID = 0
		
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
							@iBaseTableID,
							@iBaseRecordID,
							@iParent1TableID,
							@iParent1RecordID,
							@iParent2TableID,
							@iParent2RecordID,
							@iRequiredTableID,
							@iRequiredRecordID	OUTPUT
		
						SET @iRecordID = @iRequiredRecordID
		
						IF @iRecordID > 0 
						BEGIN
							EXEC [dbo].[spASRWorkflowValidTableRecord]
								@iRequiredTableID,
								@iRecordID,
								@fValidRecordID	OUTPUT
						END
		
						IF @fValidRecordID = 0
						BEGIN
							IF @iDBRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
							BEGIN
								SELECT @iCount = COUNT(*)
								FROM ASRSysWorkflowQueueColumns QC
								INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
								WHERE WFQ.instanceID = @piInstanceID
									AND QC.columnID = @iDBColumnID
		
								IF @iCount = 1
								BEGIN
									SELECT @sValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
									FROM ASRSysWorkflowQueueColumns QC
									INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
									WHERE WFQ.instanceID = @piInstanceID
										AND QC.columnID = @iDBColumnID
		
									SET @fValidRecordID = 1
									SET @fDeletedValue = 1
								END
							END
							ELSE
							BEGIN
								IF @iDBRecord = 1
								BEGIN
									SELECT @iCount = COUNT(*)
									FROM ASRSysWorkflowInstanceValues IV
									WHERE IV.instanceID = @piInstanceID
										AND IV.columnID = @iDBColumnID
										AND IV.elementID = @iTempElementID
		
									IF @iCount = 1
									BEGIN
										SELECT @sValue = rtrim(ltrim(isnull(IV.value , '''')))
										FROM ASRSysWorkflowInstanceValues IV
										WHERE IV.instanceID = @piInstanceID
											AND IV.columnID = @iDBColumnID
											AND IV.elementID = @iTempElementID
		
										SET @fValidRecordID = 1
										SET @fDeletedValue = 1
									END
								END
							END
						END
		
						IF @fValidRecordID = 0
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC [dbo].[spASRWorkflowActionFailed] @piInstanceID, @piElementID, ''Web Form item record has been deleted or not selected.''
										
							SET @psErrorMessage = ''Error loading web form. Web Form item record has been deleted or not selected.''
							RETURN
						END
					END
						
					IF @fDeletedValue = 0
					BEGIN
						IF @iDBColumnDataType = 11 -- Date column, need to format into MM\DD\YYYY
						BEGIN
							SET @sSQL = ''SELECT @sValue = convert(varchar(100), '' + @sTableName + ''.'' + @sColumnName + '', 101)''
						END
						ELSE
						BEGIN
							SET @sSQL = ''SELECT @sValue = '' + @sTableName + ''.'' + @sColumnName
						END
						
						SET @sSQL = @sSQL +
								'' FROM '' + @sTableName +
								'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(100), @iRecordID)
						SET @sSQLParam = N''@sValue varchar(MAX) OUTPUT''
						EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT
					END
				END
		
				IF @iItemType = 4
				BEGIN
					-- Workflow value.
					SELECT @sValue = ASRSysWorkflowInstanceValues.value, 
						@iType = ASRSysWorkflowElementItems.itemType,
						@iColumnID = ASRSysWorkflowElementItems.lookupColumnID
					FROM ASRSysWorkflowInstanceValues
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
					INNER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.ElementID
					WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
						AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
						AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowElementItems.identifier = @sWFValueIdentifier
		
					IF @iType = 14 -- Lookup, need to get the column data type
					BEGIN
						SELECT @iType = 
							CASE
								WHEN ASRSysColumns.dataType = -7 THEN 6 -- Logic
								WHEN ASRSysColumns.dataType = 2 THEN 5 -- Numeric
								WHEN ASRSysColumns.dataType = 4 THEN 5 -- Integer
								WHEN ASRSysColumns.dataType = 11 THEN 7 -- Date
								ELSE 3
							END
						FROM ASRSysColumns
						WHERE ASRSysColumns.columnID = @iColumnID
					END
				END
		
				IF @iItemType = 2 
				BEGIN
					-- Label with calculated caption
					EXEC [dbo].[spASRSysWorkflowCalculation]
						@piInstanceID,
						@iCalcID,
						@iResultType OUTPUT,
						@sResult OUTPUT,
						@fResult OUTPUT,
						@dtResult OUTPUT,
						@fltResult OUTPUT, 
						0
		
					SET @sValue = @sResult
					SET @iType = 3 -- Character
				END
		
				IF (@iItemType = 3)
					OR (@iItemType = 5)
					OR (@iItemType = 6)
					OR (@iItemType = 7)
					OR (@iItemType = 11)
					OR (@iItemType = 17)
				BEGIN
					IF @iStatus = 7 -- Previously SavedForLater
					BEGIN
						SELECT @sValue = 
							CASE
								WHEN (@iItemType = 6 AND IVs.value = ''1'') THEN ''TRUE'' 
								WHEN (@iItemType = 6 AND IVs.value <> ''1'') THEN ''FALSE'' 
								WHEN (@iItemType = 7 AND (upper(ltrim(rtrim(IVs.value))) = ''NULL'')) THEN '''' 
								WHEN (@iItemType = 17 AND IVs.fileUpload_File IS null) THEN ''0''
								WHEN (@iItemType = 17 AND NOT IVs.fileUpload_File IS null) THEN ''1''
								ELSE isnull(IVs.value, '''')
							END
						FROM ASRSysWorkflowInstanceValues IVs
						WHERE IVs.instanceID = @piInstanceID
							AND IVs.elementID = @piElementID
							AND IVs.identifier = @sIdentifier
					END
					ELSE	
					BEGIN
						IF @iDefaultValueType = 3 -- Calculated
						BEGIN
							EXEC [dbo].[spASRSysWorkflowCalculation]
								@piInstanceID,
								@iCalcID,
								@iResultType OUTPUT,
								@sResult OUTPUT,
								@fResult OUTPUT,
								@dtResult OUTPUT,
								@fltResult OUTPUT, 
								0
		
							IF @iItemType = 3 SET @sResult = LEFT(@sResult, @iSize)
							IF @iItemType = 5
							BEGIN
								IF @fltResult >= power(10, @iSize - @iDecimals) SET @fltResult = 0
								IF @fltResult <= (-1 * power(10, @iSize - @iDecimals)) SET @fltResult = 0
							END
		
							SET @sValue = 
								CASE
									WHEN @iResultType = 2 THEN STR(@fltResult, 100, @iDecimals)
									WHEN @iResultType = 3 THEN 
										CASE 
											WHEN @fResult = 1 THEN ''TRUE''
											ELSE ''FALSE''
										END
									WHEN @iResultType = 4 THEN convert(varchar(100), @dtResult, 101)
									ELSE convert(varchar(100), @sResult)
								END
		
							SET @iType = @iResultType
						END
						ELSE
						BEGIN
							SELECT @sValue = isnull(EIs.inputDefault, '''')
							FROM ASRSysWorkflowElementItems EIs
							WHERE EIs.elementID = @piElementID
								AND EIs.ID = @iID
						END
					END
				END		
		
				INSERT INTO @itemValues (ID, value, type)
				VALUES (@iID, @sValue, @iType)
		
				FETCH NEXT FROM itemCursor INTO 
					@iID, 
					@iItemType, 
					@iDBColumnID, 
					@iDBRecord, 
					@sWFFormIdentifier, 
					@sWFValueIdentifier, 
					@iCalcID, 
					@sIdentifier, 
					@iDefaultValueType,
					@iSize,
					@iDecimals
			END
			CLOSE itemCursor
			DEALLOCATE itemCursor
		
			SELECT thisFormItems.*, 
				IV.value, 
				IV.type AS [sourceItemType]
			FROM ASRSysWorkflowElementItems thisFormItems
			LEFT OUTER JOIN @itemValues IV ON thisFormItems.ID = IV.ID
			WHERE thisFormItems.elementID = @piElementID
			ORDER BY thisFormItems.ZOrder DESC
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRGetWorkflowQueryString
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowQueryString]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowQueryString];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowQueryString]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowQueryString]
		(
			@piInstanceID	integer,
			@piElementID	integer,
			@psQueryString	varchar(MAX)	output
		)
		AS
		BEGIN
			DECLARE
				@hResult		integer,
				@objectToken	integer,
				@sURL			varchar(MAX),
				@sParam1		varchar(MAX),
				@sDBName		sysname,
				@sSQLVersion	integer;
		
			SET @psQueryString = '''';
			SET @sSQLVersion = dbo.udfASRSQLVersion();
		
			SELECT @sURL = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_URL'';
				
			IF upper(right(@sURL, 5)) <> ''.ASPX''
				AND right(@sURL, 1) <> ''/''
				AND len(@sURL) > 0
			BEGIN
				SET @sURL = @sURL + ''/'';
			END
		
			SELECT @sParam1 = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_Web1'';
		
			IF (len(@sURL) > 0)
			BEGIN
				SET @sDBName = db_name();
		
				SELECT @psQueryString = dbo.[udfASRNetGetWorkflowQueryString]( @piInstanceID, @piElementID, @sParam1, @@servername, @sDBName);
			
				IF len(@psQueryString) > 0
				BEGIN
					SET @psQueryString = @sURL + ''?'' + @psQueryString;
				END
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRInstantiateWorkflow
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRInstantiateWorkflow]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRInstantiateWorkflow];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRInstantiateWorkflow]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'Alter PROCEDURE [dbo].[spASRInstantiateWorkflow]
		(
			@piWorkflowID	integer,			
			@piInstanceID	integer			OUTPUT,
			@psFormElements	varchar(MAX)	OUTPUT,
			@psMessage		varchar(MAX)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iInitiatorID			integer,
				@iStepID				integer,
				@iElementID				integer,
				@iRecordID				integer,
				@iRecordCount			integer,
				@sSQL					nvarchar(MAX),
				@hResult				integer,
				@sActualLoginName		sysname,
				@fUsesInitiator			bit, 
				@iTemp					integer,
				@iStartElementID		integer,
				@iTableID				integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@sForms					varchar(MAX),
				@iCount					integer,
				@iSQLVersion			integer,
				@fExternallyInitiated	bit,
				@fEnabled				bit,
				@iElementType			integer,
				@fStoredDataOK			bit, 
				@sStoredDataMsg			varchar(MAX), 
				@sStoredDataSQL			varchar(MAX), 
				@iStoredDataTableID		integer,
				@sStoredDataTableName	varchar(255),
				@iStoredDataAction		integer, 
				@iStoredDataRecordID	integer,
				@sStoredDataRecordDesc	varchar(MAX),
				@sSPName				varchar(255),
				@iNewRecordID			integer,
				@sEvalRecDesc			varchar(MAX),
				@iResult				integer,
				@iFailureFlows			integer,
				@fSaveForLater			bit;
		
			SELECT @iSQLVersion = convert(float,substring(@@version,charindex(''-'',@@version)+2,2));
		
			DECLARE @succeedingElements table(elementID int);
		
			SET @iInitiatorID = 0;
			SET @psFormElements = '''';
			SET @psMessage = '''';
			SET @iParent1TableID = 0;
			SET @iParent1RecordID = 0;
			SET @iParent2TableID = 0;
			SET @iParent2RecordID = 0;
		
			SELECT @fExternallyInitiated = CASE
					WHEN initiationType = 2 THEN 1
					ELSE 0
				END,
				@fEnabled = enabled
			FROM ASRSysWorkflows
			WHERE ID = @piWorkflowID;
		
			IF @fExternallyInitiated = 1
			BEGIN
				IF @fEnabled = 0
				BEGIN
					/* Workflow is disabled. */
					SET @psMessage = ''This link is currently disabled.'';
					RETURN
				END
		
				SET @sActualLoginName = ''<External>'';
			END
			ELSE
			BEGIN
				SET @sActualLoginName = SUSER_SNAME();
				
				SET @sSQL = ''spASRSysGetCurrentUserRecordID'';
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @hResult = 0;
			
					EXEC @hResult = @sSQL 
						@iRecordID OUTPUT,
						@iRecordCount OUTPUT;
				END
			
				IF NOT @iRecordID IS null SET @iInitiatorID = @iRecordID
				IF @iInitiatorID = 0 
				BEGIN
					/* Unable to determine the initiator''s record ID. Is it needed anyway? */
					EXEC [dbo].[spASRWorkflowUsesInitiator]
						@piWorkflowID,
						@fUsesInitiator OUTPUT;
				
					IF @fUsesInitiator = 1
					BEGIN
						IF @iRecordCount = 0
						BEGIN
							/* No records for the initiator. */
							SET @psMessage = ''Unable to locate your personnel record.'';
						END
						IF @iRecordCount > 1
						BEGIN
							/* More than one record for the initiator. */
							SET @psMessage = ''You have more than one personnel record.'';
						END
			
						RETURN
					END	
				END
				ELSE
				BEGIN
					SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel'';
		
					IF @iTableID = 0 
					BEGIN
						SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
						FROM ASRSysModuleSetup
						WHERE moduleKey = ''MODULE_WORKFLOW''
						AND parameterKey = ''Param_TablePersonnel'';
					END
		
					exec [dbo].[spASRGetParentDetails]
						@iTableID,
						@iInitiatorID,
						@iParent1TableID	OUTPUT,
						@iParent1RecordID	OUTPUT,
						@iParent2TableID	OUTPUT,
						@iParent2RecordID	OUTPUT;
				END
			END
		
			/* Create the Workflow Instance record, and remember the ID. */
			INSERT INTO [dbo].[ASRSysWorkflowInstances] (workflowID, 
				[initiatorID], 
				[status], 
				[userName], 
				[parent1TableID],
				[parent1RecordID],
				[parent2TableID],
				[parent2RecordID])
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID);
						
			SELECT @piInstanceID = MAX(id)
			FROM [dbo].[ASRSysWorkflowInstances];
		
			/* Create the Workflow Instance Steps records. 
			Set the first steps'' status to be 1 (pending Workflow Engine action). 
			Set all subsequent steps'' status to be 0 (on hold). */
		
			SELECT @iStartElementID = ASRSysWorkflowElements.ID
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.type = 0 -- Start element
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID;
		
			INSERT INTO @succeedingElements 
				SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0);
		
			INSERT INTO [dbo].[ASRSysWorkflowInstanceSteps] (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
			SELECT 
				@piInstanceID, 
				ASRSysWorkflowElements.ID, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 3
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN 1
					ELSE 0
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					WHEN ASRSysWorkflowElements.ID IN (SELECT suc.elementID
						FROM @succeedingElements suc) THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
					ELSE null
				END, 
				CASE
					WHEN ASRSysWorkflowElements.type = 0 THEN 1
					ELSE 0
				END,
				0,
				0
			FROM ASRSysWorkflowElements 
			WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID;
		
			/* Create the Workflow Instance Value records. */
			INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
			SELECT @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElementItems.identifier
			FROM ASRSysWorkflowElementItems 
			INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 2
				AND (ASRSysWorkflowElementItems.itemType = 3 
					OR ASRSysWorkflowElementItems.itemType = 5
					OR ASRSysWorkflowElementItems.itemType = 6
					OR ASRSysWorkflowElementItems.itemType = 7
					OR ASRSysWorkflowElementItems.itemType = 11
					OR ASRSysWorkflowElementItems.itemType = 13
					OR ASRSysWorkflowElementItems.itemType = 14
					OR ASRSysWorkflowElementItems.itemType = 15
					OR ASRSysWorkflowElementItems.itemType = 17
					OR ASRSysWorkflowElementItems.itemType = 0)
			UNION
			SELECT  @piInstanceID, ASRSysWorkflowElements.ID, 
				ASRSysWorkflowElements.identifier
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowElements.type = 5;
						
			SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
					
			WHILE @iCount > 0 
			BEGIN
				DECLARE immediateSubmitCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowInstanceSteps.elementID, 
					ASRSysWorkflowElements.type
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
				OPEN immediateSubmitCursor;
				FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				WHILE (@@fetch_status = 0) 
				BEGIN
					IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
					BEGIN
						SET @fStoredDataOK = 1;
						SET @sStoredDataMsg = '''';
						SET @sStoredDataRecordDesc = '''';
		
						EXEC [spASRGetStoredDataActionDetails]
							@piInstanceID,
							@iElementID,
							@sStoredDataSQL			OUTPUT, 
							@iStoredDataTableID		OUTPUT,
							@sStoredDataTableName	OUTPUT,
							@iStoredDataAction		OUTPUT, 
							@iStoredDataRecordID	OUTPUT;
		
						IF @iStoredDataAction = 0 -- Insert
						BEGIN
							SET @sSPName  = ''spASRWorkflowInsertNewRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iNewRecordID  OUTPUT, 
									@iStoredDataTableID,
									@sStoredDataSQL;
		
								SET @iStoredDataRecordID = @iNewRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 1 -- Update
						BEGIN
							SET @sSPName  = ''spASRWorkflowUpdateRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataSQL,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE IF @iStoredDataAction = 2 -- Delete
						BEGIN
							EXEC [dbo].[spASRRecordDescription]
								@iStoredDataTableID,
								@iStoredDataRecordID,
								@sStoredDataRecordDesc OUTPUT;
		
							SET @sSPName  = ''spASRWorkflowDeleteRecord'';
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID;
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0;
								SET @sStoredDataMsg = ERROR_MESSAGE();
							END CATCH
						END
						ELSE
						BEGIN
							SET @fStoredDataOK = 0;
							SET @sStoredDataMsg = ''Unrecognised data action.'';
						END
		
						IF (@fStoredDataOK = 1)
							AND ((@iStoredDataAction = 0)
								OR (@iStoredDataAction = 1))
						BEGIN
		
							EXEC [dbo].[spASRStoredDataFileActions]
								@piInstanceID,
								@iElementID,
								@iStoredDataRecordID;
						END
		
						IF @fStoredDataOK = 1
						BEGIN
							SET @sStoredDataMsg = ''Successfully '' +
								CASE
									WHEN @iStoredDataAction = 0 THEN ''inserted''
									WHEN @iStoredDataAction = 1 THEN ''updated''
									ELSE ''deleted''
								END + '' record'';
		
							IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
							BEGIN
								IF @iStoredDataRecordID > 0 
								BEGIN	
									EXEC [dbo].[spASRRecordDescription] 
										@iStoredDataTableID,
										@iStoredDataRecordID,
										@sEvalRecDesc OUTPUT;
									IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc;
								END
							END
		
							IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')'';
		
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = convert(varchar(MAX), @iStoredDataRecordID), 
								ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @iElementID
								AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
								AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
		
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 3,
								ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.message = @sStoredDataMsg
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
		
							-- Get this immediate element''s succeeding elements
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 1
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
						END
						ELSE
						BEGIN
							-- Check if the failed element has an outbound flow for failures.
							SELECT @iFailureFlows = COUNT(*)
							FROM ASRSysWorkflowElements Es
							INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
								AND Ls.startOutboundFlowCode = 1
							WHERE Es.ID = @iElementID
								AND Es.type = 5; -- 5 = StoredData
		
							IF @iFailureFlows = 0
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 4,	-- 4 = failed
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID;
		
								UPDATE ASRSysWorkflowInstances
								SET status = 2	-- 2 = error
								WHERE ID = @piInstanceID;
		
								SET @psMessage = @sStoredDataMsg;
								RETURN;
							END
							ELSE
							BEGIN
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
								SET [Status] = 8,	-- 8 = failed action
									[Message] = @sStoredDataMsg,
									[failedCount] = isnull(failedCount, 0) + 1,
									[completionCount] = isnull(completionCount, 0) - 1
								WHERE [instanceID] = @piInstanceID
									AND [elementID] = @iElementID;
		
								UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
									SET ASRSysWorkflowInstanceSteps.status = 1
									WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
										AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC);
							END
						END
					END
					ELSE
					BEGIN
						EXEC [dbo].[spASRSubmitWorkflowStep] 
							@piInstanceID, 
							@iElementID, 
							'''', 
							@sForms OUTPUT, 
							@fSaveForLater OUTPUT;
					END
		
					FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, @iElementType;
				END
				CLOSE immediateSubmitCursor;
				DEALLOCATE immediateSubmitCursor;
		
				SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
					FROM [dbo].[ASRSysWorkflowInstanceSteps]
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND (ASRSysWorkflowElements.type = 4 
							OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
							OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;
			END						
		
			/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
			DECLARE @succeedingSteps table(stepID int)
			
			INSERT INTO @succeedingSteps 
				(stepID) VALUES (-1)
		
			DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysWorkflowInstanceSteps.ID,
				ASRSysWorkflowInstanceSteps.elementID
			FROM [dbo].[ASRSysWorkflowInstanceSteps]
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE (ASRSysWorkflowInstanceSteps.status = 1 OR ASRSysWorkflowInstanceSteps.status = 2)
				AND ASRSysWorkflowElements.type = 2
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID;	
		
			OPEN formsCursor;
			FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			WHILE (@@fetch_status = 0) 
			BEGIN
				SET @psFormElements = @psFormElements + convert(varchar(MAX), @iElementID) + char(9);
		
				INSERT INTO @succeedingSteps 
				(stepID) VALUES (@iStepID)
		
				FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
			END
		
			CLOSE formsCursor;
			DEALLOCATE formsCursor;
		
			UPDATE [dbo].[ASRSysWorkflowInstanceSteps]
			SET ASRSysWorkflowInstanceSteps.status = 2, 
				userName = @sActualLoginName
			WHERE ASRSysWorkflowInstanceSteps.ID IN (SELECT stepID FROM @succeedingSteps)
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRMaternityExpectedReturn
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRMaternityExpectedReturn]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRMaternityExpectedReturn];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRMaternityExpectedReturn]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRMaternityExpectedReturn] (
			@pdblResult datetime OUTPUT,
			@EWCDate datetime,
			@LeaveStart datetime,
			@BabyBirthDate datetime,
			@Ordinary varchar(MAX)
			)
		AS
		BEGIN
		
			IF LOWER(@Ordinary) = ''ordinary''
				IF DateDiff(d,''04/06/2003'', @EWCDate) >= 0
					SET @pdblResult = Dateadd(ww,26,@LeaveStart);
				ELSE
					IF DateDiff(d,''04/30/2000'', @EWCDate) >= 0
						SET @pdblResult = Dateadd(ww,18,@LeaveStart);
					ELSE
						SET @pdblResult = Dateadd(ww,14,@LeaveStart);
			ELSE
				IF DateDiff(d,''04/06/2003'', @EWCDate) >= 0
					SET @pdblResult = Dateadd(ww,52,@LeaveStart);
				ELSE
					--29 weeks from baby birth date (but return on the monday before!)
					SET @pdblResult = DateAdd(d,203 - datepart(dw,DateAdd(d,-2,@BabyBirthDate)),@BabyBirthDate);
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRNetOutlookBatch
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRNetOutlookBatch]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRNetOutlookBatch];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRNetOutlookBatch]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRNetOutlookBatch]
		(
			@Content			varchar(MAX)	OUTPUT,
			@AllDayEvent		bit OUTPUT,
			@StartDate			datetime		OUTPUT,
			@EndDate			datetime		OUTPUT,
			@StartTime			varchar(100)	OUTPUT,
			@EndTime			varchar(100)	OUTPUT,
			@Subject			varchar(MAX)	OUTPUT,
			@Folder				varchar(MAX)	OUTPUT,
			@LinkID				integer,
			@RecordID			integer,
			@FolderID			integer,
			@StartDateColumnID	integer,
			@EndDateColumnID	integer,
			@FixedStartTime		varchar(100),
			@FixedEndTime		varchar(100),
			@StartTimeColumnID	integer,
			@EndTimeColumnID	integer,
			@TimeRange			integer,
			@Title				varchar(MAX),
			@SubjectExprID		integer,
			@RecordDescExprID	integer,
			@DateFormat			varchar(100),
			@FolderPath			varchar(MAX),
			@FolderType			integer,
			@FolderExprID		integer)
		AS
		BEGIN
		
			DECLARE @sSQL nvarchar(MAX);
			DECLARE @sParamDefinition nvarchar(500);
			DECLARE @CharValue varchar(MAX);
			DECLARE @Heading varchar(MAX);
			DECLARE @TableName varchar(MAX);
			DECLARE @ColumnName varchar(MAX);
			DECLARE @DataType integer;
				
			SELECT @sSQL = ''SELECT @StartDate=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(100),@RecordID)
			FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
			WHERE ColumnID = @StartDateColumnID;
			SET @sParamDefinition = N''@StartDate datetime OUTPUT'';
			EXEC sp_executesql @sSQL,  @sParamDefinition, @StartDate OUTPUT;
		
			SET @EndDate = Null
			IF @EndDateColumnID > 0
			BEGIN
				SELECT @sSQL = ''SELECT @EndDate=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(100),@RecordID)
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
		
				SELECT @sSQL = ''SELECT @StartTime=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(100),@RecordID)
				FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
				WHERE ColumnID = @StartTimeColumnID
				SET @sParamDefinition = N''@StartTime varchar(100) OUTPUT''
				EXEC sp_executesql @sSQL,  @sParamDefinition, @StartTime OUTPUT
		
				SELECT @sSQL = ''SELECT @EndTime=[''+ColumnName+''] FROM [''+TableName+''] WHERE ID = ''+convert(nvarchar(100),@RecordID)
				FROM ASRSysColumns JOIN ASRSysTables ON ASRSysColumns.TableID = ASRSysTables.TableID
				WHERE ColumnID = @EndTimeColumnID
				SET @sParamDefinition = N''@EndTime varchar(100) OUTPUT''
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
					IF EXISTS(SELECT * FROM sysobjects WHERE type = ''''P'''' AND name = ''''sp_ASRExpr_''+convert(nvarchar(100),@SubjectExprID)+'''''')
				             BEGIN
				                EXEC @hResult = sp_ASRExpr_''+convert(nvarchar(100),@SubjectExprID)+'' @Subject OUTPUT, ''+convert(nvarchar(100),@RecordID)+''
				                IF @hResult <> 0 SET @Subject = ''''''''
				                SET @Subject = CONVERT(varchar(255), @Subject)
					     END
					     ELSE SET @Subject = ''''''''''
				SET @sParamDefinition = N''@Subject varchar(MAX) OUTPUT''
				EXEC sp_executesql @sSQL,  @sParamDefinition, @Subject OUTPUT
			END
			ELSE
			BEGIN
				IF @RecordDescExprID > 0
				BEGIN
					SET @sSQL = ''DECLARE @hResult int
						IF EXISTS(SELECT * FROM sysobjects WHERE type = ''''P'''' AND name = ''''sp_ASRExpr_''+convert(nvarchar(100),@RecordDescExprID)+'''''')
					             BEGIN
					                EXEC @hResult = sp_ASRExpr_''+convert(nvarchar(100),@RecordDescExprID)+'' @Subject OUTPUT, ''+convert(nvarchar(100),@RecordID)+''
					                IF @hResult <> 0 SET @Subject = ''''''''
					                SET @Subject = CONVERT(varchar(255), @Subject)
						     END
						     ELSE SET @Subject = ''''''''''
					SET @sParamDefinition = N''@Subject varchar(MAX) OUTPUT''
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
					IF EXISTS(SELECT * FROM sysobjects WHERE type = ''''P'''' AND name = ''''sp_ASRExpr_''+convert(nvarchar(100),@FolderExprID)+'''''')
				             BEGIN
				                EXEC @hResult = sp_ASRExpr_''+convert(nvarchar(100),@FolderExprID)+'' @Folder OUTPUT, ''+convert(nvarchar(100),@RecordID)+''
				                IF @hResult <> 0 SET @Folder = ''''''''
					     END
					     ELSE SET @Folder = ''''''''''
				SET @sParamDefinition = N''@Folder varchar(MAX) OUTPUT''
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
					SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+isnull([''+@ColumnName+''],'''''''') FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(100),@RecordID)
				IF @DataType = 11
					SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+case when [''+@ColumnName+''] is null then ''''<Empty>'''' else convert(varchar(255),[''+@ColumnName+''],''+@DateFormat+'') end FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(100),@RecordID)
				IF @DataType = -7
					SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+case when [''+@ColumnName+''] = 1 then ''''Y'''' else ''''N'''' end FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(100),@RecordID)
				IF @DataType <> 11 AND @DataType <> 12 AND @DataType <> -7
					SELECT @sSQL = ''SELECT @CharValue=''''''+@Heading+''''''+convert(varchar(255),isnull([''+@ColumnName+''],'''''''')) FROM [''+@TableName+''] WHERE ID = ''+convert(nvarchar(100),@RecordID)
		
				SET @sParamDefinition = N''@CharValue varchar(MAX) OUTPUT''
				EXEC sp_executesql @sSQL,  @sParamDefinition, @CharValue OUTPUT
		
				IF @CharValue IS Null SET @CharValue = ''''
				SET @Content = @CharValue + char(13) + @Content
		
				FETCH NEXT FROM CursorColumns
				INTO	@Heading, @TableName, @ColumnName, @DataType
			END
		
			CLOSE CursorColumns
			DEALLOCATE CursorColumns
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRParentalLeaveEntitlement
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRParentalLeaveEntitlement]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRParentalLeaveEntitlement];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRParentalLeaveEntitlement]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRParentalLeaveEntitlement] (
			@pdblResult		float OUTPUT,
			@DateOfBirth	datetime,
			@AdoptedDate	datetime,
			@Disabled		bit,
			@Region			varchar(MAX)
		)
		AS
		BEGIN
		
			DECLARE @Today datetime,
				@ChildAge int,
				@Adopted bit,
				@YearsOfResponsibility int,
				@StartDate datetime,
				@Standard int,
				@Extended int;
		
			SET @Standard = 65;
			SET @Extended = 90;
			IF @Region = ''Rep of Ireland''
			BEGIN
				SET @Standard = 70;
				SET @Extended = 70;
			END
		
		
			--Check if we should used the Date of Birth or the Date of Adoption column...
			SET @Adopted = 0;
			SET @StartDate = @DateOfBirth;
			IF NOT @AdoptedDate IS NULL
			BEGIN
				SET @Adopted = 1;
				SET @StartDate = @AdoptedDate;
			END
		
			--Set variables based on this date...
			--(years of responsibility = years since born or adopted)
			SET @Today = getdate();
			EXEC [dbo].[sp_ASRFn_WholeYearsBetweenTwoDates] @ChildAge OUTPUT, @DateOfBirth, @Today;
			EXEC [dbo].[sp_ASRFn_WholeYearsBetweenTwoDates] @YearsOfResponsibility OUTPUT, @StartDate, @Today;
		
			SELECT @pdblResult = CASE
				WHEN @Disabled = 0 And @Adopted = 0 And @ChildAge < 5
					THEN @Standard
				WHEN @Disabled = 0 And @Adopted = 1 And @ChildAge < 18
					And @YearsOfResponsibility < 5 THEN	@Standard
				WHEN @Disabled = 1 And @Adopted = 0 And @ChildAge < 18 
					And DateDiff(d,''12/15/1994'',@DateOfBirth) >= 0 THEN	@Extended
				WHEN @Disabled = 1 And @Adopted = 1 And @ChildAge < 18 
				And DateDiff(d,''12/15/1994'',@AdoptedDate) >= 0 THEN	@Extended
				ELSE
					0
				END;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRRecordDescription
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRRecordDescription]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRRecordDescription];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRRecordDescription]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRRecordDescription]
		(
			@piTableID				integer,
			@piRecordID				integer,
			@psRecordDescription	varchar(MAX)	OUTPUT
		)
		 AS
		BEGIN
			DECLARE @sSQL varchar(MAX),
				@iRecordDescID integer,
				@sRecordDesc varchar(MAX);
		
			SET @psRecordDescription = '''';
		
			SELECT @iRecordDescID = ISNULL(ASRSysTables.recordDescExprID, 0)
				FROM ASRSysTables
				WHERE ASRSysTables.tableID = @piTableID;
		
			IF @iRecordDescID > 0 
			BEGIN
				SET @sSQL = ''sp_ASRExpr_'' + convert(varchar,@iRecordDescID);
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					EXEC @sSQL @sRecordDesc OUTPUT, @piRecordID;
					SET @psRecordDescription = @sRecordDesc;
				END
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRSendMail
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSendMail]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSendMail];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRSendMail]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRSendMail](
			@hResult int OUTPUT,
			@To varchar(MAX),
			@CC varchar(MAX),
			@BCC varchar(MAX),
			@Subject varchar(MAX),
			@Message varchar(MAX),
			@Attachment varchar(MAX))
		AS
		BEGIN
			EXEC @hResult = master..xp_sendmail
				@recipients=@To,
				@copy_recipients=@CC,
				@blind_copy_recipients=@BCC,
				@subject=@Subject,
				@message=@Message,
				@attachments=@Attachment;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRSubmitWorkflowStep
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSubmitWorkflowStep]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSubmitWorkflowStep];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRSubmitWorkflowStep]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRSubmitWorkflowStep]
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@psFormInput1		varchar(MAX),
			@psFormElements		varchar(MAX)	OUTPUT,
			@pfSavedForLater	bit				OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iIndex1			integer,
				@iIndex2			integer,
				@iID				integer,
				@sID				varchar(MAX),
				@sValue				varchar(MAX),
				@iElementType		integer,
				@iPreviousElementID	integer,
				@iValue				integer,
				@hResult			integer,
				@hTmpResult			integer,
				@sTo				varchar(MAX),
				@sCopyTo			varchar(MAX),
				@sTempTo			varchar(MAX),
				@sMessage			varchar(MAX),
				@sMessage_HypertextLinks	varchar(MAX),
				@sHypertextLinkedSteps		varchar(MAX),
				@iEmailID			integer,
				@iEmailCopyID		integer,
				@iTempEmailID		integer,
				@iEmailLoop			integer,
				@iEmailRecord		integer,
				@iEmailRecordID		integer,
				@sSQL				nvarchar(MAX),
				@iCount				integer,
				@superCursor		cursor,
				@curDelegatedRecords	cursor,
				@fDelegate			bit,
				@fDelegationValid	bit,
				@iDelegateEmailID	integer,
				@iDelegateRecordID	integer,
				@sTemp				varchar(MAX),
				@sDelegateTo		varchar(MAX),
				@sAllDelegateTo		varchar(MAX),
				@iCurrentStepID		int,
				@sDelegatedMessage	varchar(MAX),
				@iTemp				integer, 
				@iPrevElementType	integer,
				@iWorkflowID		integer,
				@sRecSelIdentifier	varchar(MAX),
				@sRecSelWebFormIdentifier	varchar(MAX), 
				@iStepID			int,
				@iElementID			int,
				@sUserName			varchar(MAX),
				@sUserEmail			varchar(MAX), 
				@sValueDescription	varchar(MAX),
				@iTableID			integer,
				@iRecDescID			integer,
				@sEvalRecDesc		varchar(MAX),
				@sExecString		nvarchar(MAX),
				@sParamDefinition	nvarchar(500),
				@sIdentifier		varchar(MAX),
				@iItemType			integer,
				@iDataAction		integer, 
				@fValidRecordID		bit,
				@iEmailTableID		integer,
				@iEmailType			integer,
				@iBaseTableID		integer,
				@iBaseRecordID		integer,
				@iRequiredRecordID	integer,
				@iParent1TableID	int,
				@iParent1RecordID	int,
				@iParent2TableID	int,
				@iParent2RecordID	int,
				@iTempElementID		integer,
				@iTrueFlowType		integer,
				@iExprID			integer,
				@iResultType		integer,
				@sResult			varchar(MAX),
				@fResult			bit,
				@dtResult			datetime,
				@fltResult			float,
				@sEmailSubject		varchar(200),
				@iTempID			integer,
				@iBehaviour			integer;
		
			SET @pfSavedForLater = 0;
		
			SELECT @iCurrentStepID = ID
			FROM ASRSysWorkflowInstanceSteps
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
		
			SET @iDelegateEmailID = 0;
			SELECT @sTemp = ISNULL(parameterValue, '''')
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_DelegateEmail'';
			SET @iDelegateEmailID = convert(integer, @sTemp);
		
			SET @psFormElements = '''';
						
			-- Get the type of the given element 
			SELECT @iElementType = E.type,
				@iEmailID = E.emailID,
				@iEmailCopyID = isnull(E.emailCCID, 0),
				@iEmailRecord = E.emailRecord, 
				@iWorkflowID = E.workflowID,
				@sRecSelIdentifier = E.RecSelIdentifier, 
				@sRecSelWebFormIdentifier = E.RecSelWebFormIdentifier, 
				@iTableID = E.dataTableID,
				@iDataAction = E.dataAction, 
				@iTrueFlowType = isnull(E.trueFlowType, 0), 
				@iExprID = isnull(E.trueFlowExprID, 0), 
				@sEmailSubject = ISNULL(E.emailSubject, '''')
			FROM ASRSysWorkflowElements E
			WHERE E.ID = @piElementID;
		
			--------------------------------------------------
			-- Read the submitted webForm/storedData values
			--------------------------------------------------
			IF @iElementType = 5 -- Stored Data element
			BEGIN
				SET @sValue = @psFormInput1;
				SET @sValueDescription = '''';
				SET @sMessage = ''Successfully '' +
					CASE
						WHEN @iDataAction = 0 THEN ''inserted''
						WHEN @iDataAction = 1 THEN ''updated''
						ELSE ''deleted''
					END + '' record'';
		
				IF @iDataAction = 2 -- Deleted - Record Description calculated before the record was deleted.
				BEGIN
					SET @sValueDescription = @psFormInput1;
				END
				ELSE
				BEGIN
					SET @iTemp = convert(integer, @sValue);
					IF @iTemp > 0 
					BEGIN	
						EXEC [dbo].[spASRRecordDescription] 
							@iTableID,
							@iTemp,
							@sEvalRecDesc OUTPUT
						IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc;
					END
				END
		
				IF len(@sValueDescription) > 0 SET @sMessage = @sMessage + '' ('' + @sValueDescription + '')'';
		
				UPDATE ASRSysWorkflowInstanceValues
				SET ASRSysWorkflowInstanceValues.value = @sValue, 
					ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription
				WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceValues.elementID = @piElementID
					AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
					AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
			END
			ELSE
			BEGIN
				-- Put the submitted form values into the ASRSysWorkflowInstanceValues table. 
				WHILE (charindex(CHAR(9), @psFormInput1) > 0)
				BEGIN
				
					SET @iIndex1 = charindex(CHAR(9), @psFormInput1);
					SET @iIndex2 = charindex(CHAR(9), @psFormInput1, @iIndex1+1);
					SET @sID = replace(LEFT(@psFormInput1, @iIndex1-1), '''''''', '''''''''''');
					SET @sValue = SUBSTRING(@psFormInput1, @iIndex1+1, @iIndex2-@iIndex1-1);
					SET @psFormInput1 = SUBSTRING(@psFormInput1, @iIndex2+1, LEN(@psFormInput1) - @iIndex2);
		
					--Get the record description (for RecordSelectors only)
					SET @sValueDescription = '''';
		
					-- Get the WebForm item type, etc.
					SELECT @sIdentifier = EI.identifier,
						@iItemType = EI.itemType,
						@iTableID = EI.tableID,
						@iBehaviour = EI.behaviour
					FROM ASRSysWorkflowElementItems EI
					WHERE EI.ID = convert(integer, @sID);
		
					SET @iParent1TableID = 0;
					SET @iParent1RecordID = 0;
					SET @iParent2TableID = 0;
					SET @iParent2RecordID = 0;
		
					IF @iItemType = 11 -- Record Selector
					BEGIN
						-- Get the table record description ID. 
						SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
						FROM ASRSysTables 
						WHERE ASRSysTables.tableID = @iTableID;
		
						SET @iTemp = convert(integer, isnull(@sValue, ''0''));
		
						-- Get the record description. 
						IF (NOT @iRecDescID IS null) AND (@iRecDescID > 0) AND (@iTemp > 0)
						BEGIN
							SET @sExecString = ''exec sp_ASRExpr_'' + convert(nvarchar(MAX), @iRecDescID) + '' @recDesc OUTPUT, @recID'';
							SET @sParamDefinition = N''@recDesc varchar(MAX) OUTPUT, @recID integer'';
							EXEC sp_executesql @sExecString, @sParamDefinition, @sEvalRecDesc OUTPUT, @iTemp;
							IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc;
						END
		
						-- Record the selected record''s parent details.
						exec [dbo].[spASRGetParentDetails]
							@iTableID,
							@iTemp,
							@iParent1TableID	OUTPUT,
							@iParent1RecordID	OUTPUT,
							@iParent2TableID	OUTPUT,
							@iParent2RecordID	OUTPUT;
					END
					ELSE
					IF (@iItemType = 0) and (@iBehaviour = 1) AND (@sValue = ''1'')-- SaveForLater Button
					BEGIN
						SET @pfSavedForLater = 1;
					END
		
					IF (@iItemType = 17) -- FileUpload Control
					BEGIN
						UPDATE ASRSysWorkflowInstanceValues
						SET ASRSysWorkflowInstanceValues.fileUpload_File = 
							CASE 
								WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_File
								ELSE null
							END,
							ASRSysWorkflowInstanceValues.fileUpload_ContentType = 
							CASE 
								WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_ContentType
								ELSE null
							END,
							ASRSysWorkflowInstanceValues.fileUpload_FileName = 
							CASE 
								WHEN @sValue = ''1'' THEN ASRSysWorkflowInstanceValues.tempFileUpload_FileName
								ELSE null
							END
						WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceValues.elementID = @piElementID
							AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
					END
					ELSE
					BEGIN
						UPDATE ASRSysWorkflowInstanceValues
						SET ASRSysWorkflowInstanceValues.value = @sValue, 
							ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription,
							ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID,
							ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
							ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID,
							ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
						WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceValues.elementID = @piElementID
							AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
					END
				END
		
				IF @pfSavedForLater = 1
				BEGIN
					/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 7
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
		
					RETURN;
				END
			END
					
			SET @hResult = 0;
			SET @sTo = '''';
			SET @sCopyTo = '''';
		
			--------------------------------------------------
			-- Process email element
			--------------------------------------------------
			IF @iElementType = 3 -- Email element
			BEGIN
				-- Get the email recipient. 
				SET @iEmailRecordID = 0;
				SET @sSQL = ''spASRSysEmailAddr'';
		
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @iEmailLoop = 0
					WHILE @iEmailLoop < 2
					BEGIN
						SET @hTmpResult = 0;
						SET @sTempTo = '''';
						SET @iTempEmailID = 
							CASE 
								WHEN @iEmailLoop = 1 THEN @iEmailCopyID
								ELSE isnull(@iEmailID, 0)
							END;
		
						IF @iTempEmailID > 0 
						BEGIN
							SET @fValidRecordID = 1;
		
							SELECT @iEmailTableID = isnull(tableID, 0),
								@iEmailType = isnull(type, 0)
							FROM ASRSysEmailAddress
							WHERE emailID = @iTempEmailID;
		
							IF @iEmailType = 0 
							BEGIN
								SET @iEmailRecordID = 0;
							END
							ELSE
							BEGIN
								SET @iTempElementID = 0;
		
								-- Get the record ID required. 
								IF (@iEmailRecord = 0) OR (@iEmailRecord = 4)
								BEGIN
									/* Initiator record. */
									SELECT @iEmailRecordID = ASRSysWorkflowInstances.initiatorID,
										@iParent1TableID = ASRSysWorkflowInstances.parent1TableID,
										@iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
										@iParent2TableID = ASRSysWorkflowInstances.parent2TableID,
										@iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
									FROM ASRSysWorkflowInstances
									WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
		
									SET @iBaseRecordID = @iEmailRecordID;
		
									IF @iEmailRecord = 4
									BEGIN
										-- Trigger record
										SELECT @iBaseTableID = isnull(WF.baseTable, 0)
										FROM ASRSysWorkflows WF
										INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
											AND WFI.ID = @piInstanceID;
									END
									ELSE
									BEGIN
										-- Initiator''s record
										SELECT @iBaseTableID = convert(integer, ISNULL(parameterValue, ''0''))
										FROM ASRSysModuleSetup
										WHERE moduleKey = ''MODULE_PERSONNEL''
											AND parameterKey = ''Param_TablePersonnel'';
		
										IF @iBaseTableID = 0
										BEGIN
											SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
											FROM ASRSysModuleSetup
											WHERE moduleKey = ''MODULE_WORKFLOW''
											AND parameterKey = ''Param_TablePersonnel'';
										END
									END
								END
		
								IF @iEmailRecord = 1
								BEGIN
									SELECT @iPrevElementType = ASRSysWorkflowElements.type,
										@iTempElementID = ASRSysWorkflowElements.ID
									FROM ASRSysWorkflowElements
									WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
										AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)));
		
									IF @iPrevElementType = 2
									BEGIN
										 -- WebForm
										SELECT @sValue = ISNULL(IV.value, ''0''),
											@iBaseTableID = EI.tableID,
											@iParent1TableID = IV.parent1TableID,
											@iParent1RecordID = IV.parent1RecordID,
											@iParent2TableID = IV.parent2TableID,
											@iParent2RecordID = IV.parent2RecordID
										FROM ASRSysWorkflowInstanceValues IV
										INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
										INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
										WHERE IV.instanceID = @piInstanceID
											AND IV.identifier = @sRecSelIdentifier
											AND Es.identifier = @sRecSelWebFormIdentifier
											AND Es.workflowID = @iWorkflowID
											AND IV.elementID = Es.ID;
									END
									ELSE
									BEGIN
										-- StoredData
										SELECT @sValue = ISNULL(IV.value, ''0''),
											@iBaseTableID = isnull(Es.dataTableID, 0),
											@iParent1TableID = IV.parent1TableID,
											@iParent1RecordID = IV.parent1RecordID,
											@iParent2TableID = IV.parent2TableID,
											@iParent2RecordID = IV.parent2RecordID
										FROM ASRSysWorkflowInstanceValues IV
										INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
											AND IV.identifier = Es.identifier
											AND Es.workflowID = @iWorkflowID
											AND Es.identifier = @sRecSelWebFormIdentifier
										WHERE IV.instanceID = @piInstanceID;
									END
		
									SET @iEmailRecordID = 
										CASE
											WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
											ELSE 0
										END;
		
									SET @iBaseRecordID = @iEmailRecordID;
								END
		
								SET @fValidRecordID = 1;
								IF (@iEmailRecord = 0) OR (@iEmailRecord = 1) OR (@iEmailRecord = 4)
								BEGIN
									SET @fValidRecordID = 0;
		
									EXEC [dbo].[spASRWorkflowAscendantRecordID]
										@iBaseTableID,
										@iBaseRecordID,
										@iParent1TableID,
										@iParent1RecordID,
										@iParent2TableID,
										@iParent2RecordID,
										@iEmailTableID,
										@iRequiredRecordID	OUTPUT;
		
									SET @iEmailRecordID = @iRequiredRecordID;
		
									IF @iRequiredRecordID > 0 
									BEGIN
										EXEC [dbo].[spASRWorkflowValidTableRecord]
											@iEmailTableID,
											@iEmailRecordID,
											@fValidRecordID	OUTPUT;
									END
		
									IF @fValidRecordID = 0
									BEGIN
										IF @iEmailRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
										BEGIN
											SELECT @sTempTo = rtrim(ltrim(isnull(QC.columnValue , '''')))
											FROM ASRSysWorkflowQueueColumns QC
											INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
											WHERE WFQ.instanceID = @piInstanceID
												AND QC.emailID = @iTempEmailID;
		
											IF len(@sTempTo) > 0 SET @fValidRecordID = 1;
										END
										ELSE
										BEGIN
											IF @iEmailRecord = 1
											BEGIN
												SELECT @sTempTo = rtrim(ltrim(isnull(IV.value , '''')))
												FROM ASRSysWorkflowInstanceValues IV
												WHERE IV.instanceID = @piInstanceID
													AND IV.emailID = @iTempEmailID
													AND IV.elementID = @iTempElementID;
		
												IF len(@sTempTo) > 0 SET @fValidRecordID = 1;
											END
										END
									END
		
									IF (@fValidRecordID = 0) AND (@iEmailLoop = 0)
									BEGIN
										-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
										EXEC [dbo].[spASRWorkflowActionFailed] 
											@piInstanceID, 
											@piElementID, 
											''Email record has been deleted or not selected.'';
													
										SET @hTmpResult = -1;
									END
								END
							END
		
							IF @fValidRecordID = 1
							BEGIN
								/* Get the recipient address. */
								IF len(@sTempTo) = 0
								BEGIN
									EXEC @hTmpResult = @sSQL @sTempTo OUTPUT, @iTempEmailID, @iEmailRecordID;
									IF @sTempTo IS null SET @sTempTo = '''';
								END
		
								IF (LEN(rtrim(ltrim(@sTempTo))) = 0) AND (@iEmailLoop = 0)
								BEGIN
									-- Email step failure if no known recipient.
									-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
									EXEC [dbo].[spASRWorkflowActionFailed] 
										@piInstanceID, 
										@piElementID, 
										''No email recipient.'';
												
									SET @hTmpResult = -1;
								END
							END
		
							IF @iEmailLoop = 1 
							BEGIN
								SET @sCopyTo = @sTempTo;
		
								IF (rtrim(ltrim(@sCopyTo)) = ''@'')
									OR (charindex('' @ '', @sCopyTo) > 0)
								BEGIN
									SET @sCopyTo = '''';
								END
							END
							ELSE
							BEGIN
								SET @sTo = @sTempTo;
							END
						END
						
						SET @iEmailLoop = @iEmailLoop + 1;
		
						IF @hTmpResult <> 0 SET @hResult = @hTmpResult;
					END
				END
		
				IF LEN(rtrim(ltrim(@sTo))) > 0
				BEGIN
					IF (rtrim(ltrim(@sTo)) = ''@'')
						OR (charindex('' @ '', @sTo) > 0)
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.userEmail = @sTo,
							ASRSysWorkflowInstanceSteps.emailCC = @sCopyTo
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
		
						EXEC [dbo].[spASRWorkflowActionFailed] 
							@piInstanceID, 
							@piElementID, 
							''Invalid email recipient.'';
						
						SET @hResult = -1;
					END
					ELSE
					BEGIN
						/* Build the email message. */
						EXEC [dbo].[spASRGetWorkflowEmailMessage] 
							@piInstanceID, 
							@piElementID, 
							@sMessage OUTPUT, 
							@sMessage_HypertextLinks OUTPUT, 
							@sHypertextLinkedSteps OUTPUT, 
							@fValidRecordID OUTPUT, 
							@sTo;
		
						IF @fValidRecordID = 1
						BEGIN
							exec [dbo].[spASRDelegateWorkflowEmail] 
								@sTo,
								@sCopyTo,
								@sMessage,
								@sMessage_HypertextLinks,
								@iCurrentStepID,
								@sEmailSubject;
						END
						ELSE
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC [dbo].[spASRWorkflowActionFailed] 
								@piInstanceID, 
								@piElementID, 
								''Email item database value record has been deleted or not selected.'';
										
							SET @hResult = -1;
						END
					END
				END
			END
		
			--------------------------------------------------
			-- Mark the step as complete
			--------------------------------------------------
			IF @hResult = 0
			BEGIN
				/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
					ASRSysWorkflowInstanceSteps.userEmail = CASE
						WHEN @iElementType = 3 THEN @sTo
						ELSE ASRSysWorkflowInstanceSteps.userEmail
					END,
					ASRSysWorkflowInstanceSteps.emailCC = CASE
						WHEN @iElementType = 3 THEN @sCopyTo
						ELSE ASRSysWorkflowInstanceSteps.emailCC
					END,
					ASRSysWorkflowInstanceSteps.hypertextLinkedSteps = CASE
						WHEN @iElementType = 3 THEN @sHypertextLinkedSteps
						ELSE ASRSysWorkflowInstanceSteps.hypertextLinkedSteps
					END,
					ASRSysWorkflowInstanceSteps.message = CASE
						WHEN @iElementType = 3 THEN @sMessage
						WHEN @iElementType = 5 THEN @sMessage
						ELSE ''''
					END,
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
			
				IF @iElementType = 4 -- Decision element
				BEGIN
					IF @iTrueFlowType = 1
					BEGIN
						-- Decision Element flow determined by a calculation
						EXEC [dbo].[spASRSysWorkflowCalculation]
							@piInstanceID,
							@iExprID,
							@iResultType OUTPUT,
							@sResult OUTPUT,
							@fResult OUTPUT,
							@dtResult OUTPUT,
							@fltResult OUTPUT, 
							0;
		
						SET @iValue = convert(integer, @fResult);
					END
					ELSE
					BEGIN
						-- Decision Element flow determined by a button in a preceding web form
						SET @iPrevElementType = 4; -- Decision element
						SET @iPreviousElementID = @piElementID;
		
						WHILE (@iPrevElementType = 4)
						BEGIN
							SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
								@iPrevElementType = isnull(WE.type, 0)
							FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPreviousElementID) PE
							INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
							INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
								AND WIS.instanceID = @piInstanceID;
		
							SET @iPreviousElementID = @iTempID;
						END
					
						SELECT @sValue = ISNULL(IV.value, ''0'')
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
						WHERE IV.elementID = @iPreviousElementID
							AND IV.instanceid = @piInstanceID
							AND E.ID = @piElementID;
		
						SET @iValue = 
							CASE
								WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
								ELSE 0
							END;
					END
				
					IF @iValue IS null SET @iValue = 0;
		
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
			
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 1,
						ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
						ASRSysWorkflowInstanceSteps.completionDateTime = null
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID IN 
							(SELECT SUCC.id FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, @iValue) SUCC)
						AND (ASRSysWorkflowInstanceSteps.status = 0
							OR ASRSysWorkflowInstanceSteps.status = 2
							OR ASRSysWorkflowInstanceSteps.status = 6
							OR ASRSysWorkflowInstanceSteps.status = 8
							OR ASRSysWorkflowInstanceSteps.status = 3);
				END
				ELSE
				BEGIN
					IF @iElementType <> 3 -- 3=Email element
					BEGIN
						-- Do not the following bit when the submitted element is an Email element as 
						-- the succeeding elements will already have been actioned.
						DECLARE @succeedingElements TABLE(elementID integer);
		
						EXEC [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]  
							@piInstanceID, 
							@piElementID, 
							@superCursor OUTPUT,
							'''';
		
						FETCH NEXT FROM @superCursor INTO @iTemp;
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO @succeedingElements (elementID) VALUES (@iTemp);
							
							FETCH NEXT FROM @superCursor INTO @iTemp;
						END
						CLOSE @superCursor;
						DEALLOCATE @superCursor;
		
						-- If the submitted element is a web form, then any succeeding webforms are actioned for the same user.
						IF @iElementType = 2 -- WebForm
						BEGIN
							SELECT @sUserName = isnull(WIS.userName, ''''),
								@sUserEmail = isnull(WIS.userEmail, '''')
							FROM ASRSysWorkflowInstanceSteps WIS
							WHERE WIS.instanceID = @piInstanceID
								AND WIS.elementID = @piElementID;
		
							-- Return a list of the workflow form elements that may need to be displayed to the initiator straight away 
							DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
							SELECT ASRSysWorkflowInstanceSteps.ID,
								ASRSysWorkflowInstanceSteps.elementID
							FROM ASRSysWorkflowInstanceSteps
							INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN 
									(SELECT suc.elementID
									FROM @succeedingElements suc)
								AND ASRSysWorkflowElements.type = 2
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 6
									OR ASRSysWorkflowInstanceSteps.status = 8
									OR ASRSysWorkflowInstanceSteps.status = 3);
		
							OPEN formsCursor;
							FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
							WHILE (@@fetch_status = 0) 
							BEGIN
								SET @psFormElements = @psFormElements + convert(varchar(MAX), @iElementID) + char(9);
		
								DELETE FROM ASRSysWorkflowStepDelegation
								WHERE stepID = @iStepID;
		
								INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
									(SELECT WSD.delegateEmail, @iStepID
									FROM ASRSysWorkflowStepDelegation WSD
									WHERE WSD.stepID = @iCurrentStepID);
								
								-- Change the step status to be 2 (pending user input). 
								UPDATE ASRSysWorkflowInstanceSteps
								SET ASRSysWorkflowInstanceSteps.status = 2, 
									ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
									ASRSysWorkflowInstanceSteps.completionDateTime = null,
									ASRSysWorkflowInstanceSteps.userName = @sUserName,
									ASRSysWorkflowInstanceSteps.userEmail = @sUserEmail 
								WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID
									AND (ASRSysWorkflowInstanceSteps.status = 0
										OR ASRSysWorkflowInstanceSteps.status = 2
										OR ASRSysWorkflowInstanceSteps.status = 6
										OR ASRSysWorkflowInstanceSteps.status = 8
										OR ASRSysWorkflowInstanceSteps.status = 3);
								
								FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID;
							END
							CLOSE formsCursor;
							DEALLOCATE formsCursor;
		
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 1,
								ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.completionDateTime = null
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN 
									(SELECT suc.elementID
									FROM @succeedingElements suc)
								AND ASRSysWorkflowInstanceSteps.elementID NOT IN 
									(SELECT ASRSysWorkflowElements.ID
									FROM ASRSysWorkflowElements
									WHERE ASRSysWorkflowElements.type = 2)
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 6
									OR ASRSysWorkflowInstanceSteps.status = 8
									OR ASRSysWorkflowInstanceSteps.status = 3);
						END
						ELSE
						BEGIN
							DELETE FROM ASRSysWorkflowStepDelegation
							WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
								FROM ASRSysWorkflowInstanceSteps
								WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
									AND ASRSysWorkflowInstanceSteps.elementID IN 
										(SELECT suc.elementID
										FROM @succeedingElements suc)
									AND (ASRSysWorkflowInstanceSteps.status = 0
										OR ASRSysWorkflowInstanceSteps.status = 2
										OR ASRSysWorkflowInstanceSteps.status = 6
										OR ASRSysWorkflowInstanceSteps.status = 8
										OR ASRSysWorkflowInstanceSteps.status = 3));
							
							INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
							(SELECT WSD.delegateEmail,
								SuccWIS.ID
							FROM ASRSysWorkflowStepDelegation WSD
							INNER JOIN ASRSysWorkflowInstanceSteps CurrWIS ON WSD.stepID = CurrWIS.ID
							INNER JOIN ASRSysWorkflowInstanceSteps SuccWIS ON CurrWIS.instanceID = SuccWIS.instanceID
								AND SuccWIS.elementID IN (SELECT suc.elementID
									FROM @succeedingElements suc)
								AND (SuccWIS.status = 0
									OR SuccWIS.status = 2
									OR SuccWIS.status = 6
									OR SuccWIS.status = 8
									OR SuccWIS.status = 3)
							INNER JOIN ASRSysWorkflowElements SuccWE ON SuccWIS.elementID = SuccWE.ID
								AND SuccWE.type = 2
							WHERE WSD.stepID = @iCurrentStepID);
		
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 1,
								ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.completionDateTime = null,
								ASRSysWorkflowInstanceSteps.userEmail = CASE
									WHEN (SELECT ASRSysWorkflowElements.type 
										FROM ASRSysWorkflowElements 
										WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @sTo -- 2 = Web Form element
									ELSE ASRSysWorkflowInstanceSteps.userEmail
								END
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN 
									(SELECT suc.elementID
									FROM @succeedingElements suc)
								AND (ASRSysWorkflowInstanceSteps.status = 0
									OR ASRSysWorkflowInstanceSteps.status = 2
									OR ASRSysWorkflowInstanceSteps.status = 6
									OR ASRSysWorkflowInstanceSteps.status = 8
									OR ASRSysWorkflowInstanceSteps.status = 3);
						END
					END
				END
			
				-- Set activated Web Forms to be ''pending'' (to be done by the user) 
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 2
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 2);
		
				-- Set activated Terminators to be ''completed'' 
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowElements.type = 1);
		
				-- Count how many terminators have completed. ie. if the workflow has completed. 
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type = 1;
							
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
					
					-- Steps pending action are no longer required.
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 0 -- 0 = On hold
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND (ASRSysWorkflowInstanceSteps.status = 1 -- 1 = Pending Engine Action
							OR ASRSysWorkflowInstanceSteps.status = 2); -- 2 = Pending User Action
				END
		
				IF @iElementType = 3 -- Email element
					OR @iElementType = 5 -- Stored Data element
				BEGIN
					exec [dbo].[spASREmailImmediate] ''HR Pro Workflow'';
				END
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRSysOvernightTableUpdate
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSysOvernightTableUpdate]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSysOvernightTableUpdate];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRSysOvernightTableUpdate]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRSysOvernightTableUpdate]
		(
			@piTableName varchar(255),
			@piFieldName varchar(255),
			@piBatches int
		) 
		AS
		BEGIN
			SET NOCOUNT ON;
		
			-- Create the progress table if it doesn''t already exist
			IF OBJECT_ID(''ASRSysOvernightProgress'', N''U'') IS NULL
				CREATE TABLE ASRSysOvernightProgress
					(TableName varchar(255)
					, RecCount int
					, IDRange varchar(255)
					, StartDate datetime
					, EndDate datetime
					, DurationMins int);
		
			DECLARE @lowid int,@highid int,@maxid int;
			DECLARE @rowcount int, @start datetime;
		
			DECLARE @sSQL				nvarchar(MAX),
					@sParamDefinition	nvarchar(500);
		
			-- Determine the number of ID''s we''ll update in each batch
			IF ISNULL(@piBatches, 0) = 0
				SET @piBatches = 2000;
			SET @lowid = 0 ;
			SET @highid = @lowid + @piBatches;
			
			SET @sSQL = ''SELECT @maxid = ISNULL(MAX(ID),0) FROM '' + @piTableName;
			SET @sParamDefinition = N''@maxid int OUTPUT'';
			EXEC sp_executesql @sSQL, @sParamDefinition, @maxid OUTPUT;
		
			WHILE 1=1
			BEGIN
				SET @start = GETDATE();
				
				-- Do the update
				SELECT @sSQL = ''UPDATE '' + @piTableName + '' SET '' + @piFieldName + '' = '' + @piFieldName
							+ '' WHERE ID BETWEEN @lowid AND @highid'';
				SET @sParamDefinition = N''@lowid int, @highid int'';
				EXEC sp_executesql @sSQL, @sParamDefinition, @lowid, @highid;
		
				SET @rowcount = @@ROWCOUNT;
		
				-- insert a record to this progress table to check the progress
				INSERT INTO ASRSysOvernightProgress 
					SELECT @piTableName
						, @rowcount
						, CAST(@lowid as varchar(255)) + ''-'' + CAST(@highid as varchar(255))
						, @start
						, GETDATE()
						, DATEDIFF(n, @start, GETDATE());
		
				SET @lowid = @lowid + @piBatches;
				SET @highid = @lowid + @piBatches;
		
				IF @lowid > @maxid
				BEGIN
					CHECKPOINT;
					BREAK;
				END
				ELSE
					CHECKPOINT;
			END
		
			SET NOCOUNT OFF;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRUpdateStatistics
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRUpdateStatistics]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRUpdateStatistics];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRUpdateStatistics]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRUpdateStatistics]
		AS
		BEGIN
		
			SET NOCOUNT ON;
		
			DECLARE @sTableName nvarchar(255),
					@sVarCommand nvarchar(MAX);
		
			-- Checking fragmentation
			DECLARE tables CURSOR FOR
				SELECT so.[Name]
				FROM sysobjects so
				JOIN sysindexes si ON so.id = si.id
				WHERE so.type =''U'' AND si.indid < 2 AND si.rows > 0
				ORDER BY so.[Name];
		
			-- Open the cursor
			OPEN tables;
		
			-- Loop through all the tables in the database running dbcc showcontig on each one
			FETCH NEXT FROM tables INTO @sTableName;
		
			WHILE @@FETCH_STATUS = 0
			BEGIN
				SET @sVarCommand = ''UPDATE STATISTICS ['' + @sTableName + ''] WITH FULLSCAN'';
				EXECUTE sp_executeSQL @sVarCommand;
				FETCH NEXT FROM tables INTO @sTableName;
			END
		
			-- Close and deallocate the cursor
			CLOSE tables;
			DEALLOCATE tables;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRWorkflowActionFailed
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowActionFailed]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowActionFailed];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRWorkflowActionFailed]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'Alter PROCEDURE [dbo].[spASRWorkflowActionFailed]
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@psMessage			varchar(MAX)
		)
		AS
		BEGIN
			DECLARE
				@iFailureFlows	integer,
				@iCount			integer;
		
			-- Check if the failed element has an outbound flow for failures.
			SELECT @iFailureFlows = COUNT(*)
			FROM ASRSysWorkflowElements Es
			INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
				AND Ls.startOutboundFlowCode = 1
			WHERE Es.ID = @piElementID
				AND Es.type = 5; -- 5 = StoredData
		
			IF @iFailureFlows = 0
			BEGIN
				UPDATE ASRSysWorkflowInstanceSteps
				SET status = 4,	-- 4 = failed
					message = @psMessage,
					failedCount = isnull(failedCount, 0) + 1
				WHERE instanceID = @piInstanceID
					AND elementID = @piElementID;
		
				UPDATE ASRSysWorkflowInstances
				SET status = 2	-- 2 = error
				WHERE ID = @piInstanceID;
			END
			ELSE
			BEGIN
				UPDATE ASRSysWorkflowInstanceSteps
				SET status = 8,	-- 8 = failed action
					message = @psMessage,
					failedCount = isnull(failedCount, 0) + 1
				WHERE instanceID = @piInstanceID
					AND elementID = @piElementID;
		
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 1,
					ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionDateTime = null
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID IN 
						(SELECT id 
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, 1))
					AND (ASRSysWorkflowInstanceSteps.status = 0
						OR ASRSysWorkflowInstanceSteps.status = 3
						OR ASRSysWorkflowInstanceSteps.status = 4
						OR ASRSysWorkflowInstanceSteps.status = 6
						OR ASRSysWorkflowInstanceSteps.status = 8);
								
				-- Set activated Web Forms to be ''pending'' (to be done by the user) 
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 2
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 2);
								
				-- Set activated Terminators to be ''completed''
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 1);
								
				-- Count how many terminators have completed. ie. if the workflow has completed.
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type = 1;
													
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
					
					/* NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					is performed by a DELETE trigger on the ASRSysWorkflowInstances table. */
				END
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRWorkflowFileUpload
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowFileUpload]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowFileUpload];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRWorkflowFileUpload]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRWorkflowFileUpload]
		(
			@piElementItemID	integer,
			@piInstanceID		integer,
			@pimgFile			image,
			@psContentType		varchar(MAX),
			@psFileName			varchar(MAX),
			@pfClear			bit
		)
		AS
		BEGIN
			DECLARE	@iElementID		integer,
					@sIdentifier	varchar(MAX);
		
			SELECT
				@iElementID = elementID,
				@sIdentifier = identifier
			FROM ASRSysWorkflowElementItems
			WHERE id = @piElementItemID;
		
			UPDATE ASRSysWorkflowInstanceValues 
			SET [TempFileUpload_File] = 
					CASE 
						WHEN @pfClear = 1 THEN null
						ELSE @pimgFile
					END, 
				[TempFileUpload_ContentType] = 
					CASE 
						WHEN @pfClear = 1 THEN null
						ELSE @psContentType
					END, 
				[TempFileUpload_Filename] = 
					CASE 
						WHEN @pfCLear = 1 THEN null
						ELSE @psFileName
					END
			WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceValues.elementID = @iElementID
				AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier;
		
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRWorkflowRebuild
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowRebuild]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowRebuild];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRWorkflowRebuild]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRWorkflowRebuild]
		AS
		BEGIN	
			-- Refresh all scheduled Workflow items in the queue.
			DECLARE @sTableName 	varchar(255),
				@iTableID			int,
				@sSQL				nvarchar(MAX)
			
			-- Get a cursor of the tables in the database.
			DECLARE curTables CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableName, tableID
				FROM ASRSysTables;
			OPEN curTables;
		
			DELETE FROM ASRSysWorkflowQueue 
			WHERE dateInitiated IS null 
				AND [Immediate] = 0;
		
			-- Loop through the tables in the database.
			FETCH NEXT FROM curTables INTO @sTableName, @iTableID;
			WHILE @@fetch_status <> -1
			BEGIN
				/* Get a cursor of the records in the current table.  */
				/* Call the Workflow trigger for that table and record  */
				SET @sSQL = ''
					DECLARE @iCurrentID	int,
						@sSQL		nvarchar(MAX);
					
					IF EXISTS (SELECT * FROM sysobjects
					WHERE id = object_id(''''spASRWorkflowRebuild_'' + LTrim(Str(@iTableID)) + '''''') 
						AND sysstat & 0xf = 4)
					BEGIN
						DECLARE curRecords CURSOR FOR
						SELECT id
						FROM '' + @sTableName + '';
		
						OPEN curRecords;
		
						FETCH NEXT FROM curRecords INTO @iCurrentID;
						WHILE @@fetch_status <> -1
						BEGIN
							SET @sSQL = ''''EXEC spASRWorkflowRebuild_'' + LTrim(Str(@iTableID)) 
								+ '' '''' + convert(varchar(100), @iCurrentID) + '''''''';
							EXECUTE sp_executeSQL @sSQL;
		
							FETCH NEXT FROM curRecords INTO @iCurrentID;
						END
						CLOSE curRecords;
						DEALLOCATE curRecords;
					END'';
		
				 EXECUTE sp_executeSQL @sSQL;
		
				/* Move onto the next table in the database. */ 
				FETCH NEXT FROM curTables INTO @sTableName, @iTableID;
			END
		
			CLOSE curTables;
			DEALLOCATE curTables;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRWorkflowStepDescription
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowStepDescription]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowStepDescription];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRWorkflowStepDescription]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRWorkflowStepDescription]
		(
			@piInstanceStepID	integer,
			@psDescription		varchar(MAX)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iInstanceID			integer,
				@iExprID				integer,
				@iResultType			integer,
				@sResult				varchar(MAX),
				@fResult				bit,
				@dtResult				datetime,
				@fltResult				float,
				@fDescHasWorkflowName	bit,
				@fDescHasElementCaption	bit,
				@sWorkflowName			varchar(MAX),
				@sElementCaption		varchar(MAX);
		
			-- Get the InstanceID and associated DescriptionExprID of the given step
			SELECT @iInstanceID = isnull(WIS.instanceID, 0),
				@iExprID = isnull(WEs.descriptionExprID, 0),
				@fDescHasWorkflowName = isnull(WEs.descHasWorkflowName, 0),
				@fDescHasElementCaption = isnull(WEs.descHasElementCaption, 0),
				@sWorkflowName = isnull(Ws.name, ''''),
				@sElementCaption = isnull(WEs.caption, '''')
			FROM ASRSysWorkflowInstanceSteps WIS
			INNER JOIN ASRSysWorkflowElements WEs ON WIS.elementID = WEs.ID
			INNER JOIN ASRSysWorkflows Ws ON WEs.workflowID = Ws.ID
			WHERE WIS.ID = @piInstanceStepID;
		
			IF @iExprID > 0
			BEGIN
				EXEC [dbo].[spASRSysWorkflowCalculation]
					@iInstanceID,
					@iExprID,
					@iResultType OUTPUT,
					@sResult OUTPUT,
					@fResult OUTPUT,
					@dtResult OUTPUT,
					@fltResult OUTPUT, 
					0;
			END
		
			IF @fDescHasWorkflowName = 1
			BEGIN
				SET @sResult = @sWorkflowName 
					+ '' - ''
					+ isnull(@sResult, '''');
			END
		
			IF @fDescHasElementCaption = 1
			BEGIN
				SET @sResult = @sElementCaption 
					+ '' - ''
					+ isnull(@sResult, '''');
			END
		
			SELECT @psDescription = isnull(@sResult, '''');
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRWorkflowStoredDataFile
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowStoredDataFile]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowStoredDataFile];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRWorkflowStoredDataFile]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRWorkflowStoredDataFile]
		(
			@piElementColumnID	integer,
			@piInstanceID		integer,
			@piValueType		integer			OUTPUT,
			@psFileName			varchar(MAX)	OUTPUT,
			@psErrorMessage		varchar(MAX)	OUTPUT,
			@piOLEType			integer			OUTPUT
		)
		AS
		BEGIN
			DECLARE 
				@iWorkflowID			integer,
				@iElementID				integer,
				@sElementIdentifier		varchar(MAX),
				@sItemIdentifier		varchar(MAX),
				@iDBColumnID			integer,
				@iDBRecord				integer,
				@sTableName				sysname,
				@sColumnName			sysname,
				@iRequiredTableID		integer,
				@iRequiredRecordID		integer,
				@iRecordID				integer,
				@iBaseTableID			integer,
				@iBaseRecordID			integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@iInitiatorID			integer,
				@iInitParent1TableID	integer,
				@iInitParent1RecordID	integer,
				@iInitParent2TableID	integer,
				@iInitParent2RecordID	integer,
				@iElementType			integer, 
				@iTempElementID			integer,
				@sValue					varchar(MAX),
				@fValidRecordID			bit,
				@fDeletedValue			bit,
				@iPersonnelTableID		integer,
				@iCount					integer,
				@sSQL					nvarchar(MAX),
				@sSQLParam				nvarchar(MAX);
		
			SELECT @iWorkflowID = isnull(WE.workflowID, 0),
				@iBaseTableID = isnull(WF.baseTable, 0),
				@piValueType = isnull(WEC.valueType, 0),
				@sElementIdentifier = upper(rtrim(ltrim(isnull(WEC.WFFormIdentifier, '''')))),
				@sItemIdentifier = upper(rtrim(ltrim(isnull(WEC.WFValueIdentifier, '''')))),
				@iDBColumnID = isnull(WEC.DBColumnID, 0),
				@iDBRecord = isnull(WEC.DBRecord, 0)
			FROM ASRSysWorkflowElementColumns WEC
			INNER JOIN ASRSysWorkflowElements WE ON WEC.elementID = WE.ID
			INNER JOIN ASRSysWorkflows WF ON WE.workflowID = WF.ID
			WHERE WEC.ID = @piElementColumnID;
		
			IF @piValueType = 2 -- DB File
			BEGIN
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
					@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
					@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
					@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
		
				SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel'';
		
				IF @iPersonnelTableID = 0
				BEGIN
					SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_WORKFLOW''
					AND parameterKey = ''Param_TablePersonnel'';
				END
		
				SET @fDeletedValue = 0;
		
				SELECT @sTableName = ASRSysTables.tableName, 
					@iRequiredTableID = ASRSysTables.tableID, 
					@sColumnName = ASRSysColumns.columnName,
					@piOLEType = ASRSysColumns.OLEType
				FROM ASRSysColumns
				INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
				WHERE ASRSysColumns.columnID = @iDBColumnID;
		
				IF @iDBRecord = 0
				BEGIN
					-- Initiator''s record
					SET @iRecordID = @iInitiatorID;
					SET @iParent1TableID = @iInitParent1TableID;
					SET @iParent1RecordID = @iInitParent1RecordID;
					SET @iParent2TableID = @iInitParent2TableID;
					SET @iParent2RecordID = @iInitParent2RecordID;
					SET @iBaseTableID = @iPersonnelTableID;
				END			
		
				IF @iDBRecord = 4
				BEGIN
					-- Trigger record
					SET @iRecordID = @iInitiatorID;
					SET @iParent1TableID = @iInitParent1TableID;
					SET @iParent1RecordID = @iInitParent1RecordID;
					SET @iParent2TableID = @iInitParent2TableID;
					SET @iParent2RecordID = @iInitParent2RecordID;
				END
		
				IF @iDBRecord = 1
				BEGIN
					-- Identified record.
					SELECT @iElementType = ASRSysWorkflowElements.type, 
						@iTempElementID = ASRSysWorkflowElements.ID
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
						AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sElementIdentifier)));
						
					IF @iElementType = 2
					BEGIN
						 -- WebForm
						SELECT @sValue = ISNULL(IV.value, ''0''),
							@iBaseTableID = EI.tableID,
							@iParent1TableID = IV.parent1TableID,
							@iParent1RecordID = IV.parent1RecordID,
							@iParent2TableID = IV.parent2TableID,
							@iParent2RecordID = IV.parent2RecordID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
						INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
						WHERE IV.instanceID = @piInstanceID
							AND IV.identifier = @sItemIdentifier
							AND Es.identifier = @sElementIdentifier
							AND Es.workflowID = @iWorkflowID
							AND IV.elementID = Es.ID;
					END
					ELSE
					BEGIN
						-- StoredData
						SELECT @sValue = ISNULL(IV.value, ''0''),
							@iBaseTableID = isnull(Es.dataTableID, 0),
							@iParent1TableID = IV.parent1TableID,
							@iParent1RecordID = IV.parent1RecordID,
							@iParent2TableID = IV.parent2TableID,
							@iParent2RecordID = IV.parent2RecordID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
							AND IV.identifier = Es.identifier
							AND Es.workflowID = @iWorkflowID
							AND Es.identifier = @sElementIdentifier
						WHERE IV.instanceID = @piInstanceID;
					END
		
					SET @iRecordID = 
						CASE
							WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
							ELSE 0
						END
				END	
				
				SET @iBaseRecordID = @iRecordID;
		
				IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
				BEGIN
					SET @fValidRecordID = 0;
		
					EXEC [dbo].[spASRWorkflowAscendantRecordID]
						@iBaseTableID,
						@iBaseRecordID,
						@iParent1TableID,
						@iParent1RecordID,
						@iParent2TableID,
						@iParent2RecordID,
						@iRequiredTableID,
						@iRequiredRecordID	OUTPUT;
		
					SET @iRecordID = @iRequiredRecordID;
		
					IF @iRecordID > 0 
					BEGIN
						EXEC [dbo].[spASRWorkflowValidTableRecord]
							@iRequiredTableID,
							@iRecordID,
							@fValidRecordID	OUTPUT;
					END
		
					IF @fValidRecordID = 0
					BEGIN
						IF @iDBRecord = 4 -- Trigger record. 
						BEGIN
							SELECT @iCount = COUNT(*)
							FROM ASRSysWorkflowQueueColumns QC
							INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
							WHERE WFQ.instanceID = @piInstanceID
								AND QC.columnID = @iDBColumnID;
		
							IF @iCount = 1
							BEGIN
								SELECT @sValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
								FROM ASRSysWorkflowQueueColumns QC
								INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
								WHERE WFQ.instanceID = @piInstanceID
									AND QC.columnID = @iDBColumnID;
		
								SET @fValidRecordID = 1;
								SET @fDeletedValue = 1;
							END
						END
						ELSE
						BEGIN
							IF @iDBRecord = 1
							BEGIN
								SELECT @iCount = COUNT(*)
								FROM ASRSysWorkflowInstanceValues IV
								WHERE IV.instanceID = @piInstanceID
									AND IV.columnID = @iDBColumnID
									AND IV.elementID = @iTempElementID;
		
								IF @iCount = 1
								BEGIN
									SELECT @sValue = rtrim(ltrim(isnull(IV.value , '''')))
									FROM ASRSysWorkflowInstanceValues IV
									WHERE IV.instanceID = @piInstanceID
										AND IV.columnID = @iDBColumnID
										AND IV.elementID = @iTempElementID;
		
									SET @fValidRecordID = 1;
									SET @fDeletedValue = 1;
								END
							END
						END
					END
		
					IF @fValidRecordID = 0
					BEGIN
						SET @psErrorMessage = ''Record has been deleted or not selected.'';
						RETURN;
					END
				END
					
				IF @fDeletedValue = 0
				BEGIN
					IF (@piOLEType = 0) OR (@piOLEType = 1)
					BEGIN
						SET @sSQL = ''SELECT @psFileName = '' + @sTableName + ''.'' + @sColumnName +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(255), @iRecordID);
						SET @sSQLParam = N''@psFileName varchar(MAX) OUTPUT'';
						EXEC sp_executesql @sSQL, @sSQLParam, @psFileName OUTPUT;
					END
					ELSE
					BEGIN
						SET @sSQL = ''SELECT '' + @sTableName + ''.'' + @sColumnName + '' AS [file]'' +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(255), @iRecordID);
						EXEC sp_executesql @sSQL;
					END
				END
			END
			
			IF @piValueType = 1 -- WF File
			BEGIN
				SELECT @iElementID = isnull(ID, 0)
				FROM ASRSysWorkflowElements
				WHERE workflowID = @iWorkflowID
					AND upper(ltrim(rtrim(isnull(identifier, '''')))) = @sElementIdentifier;
		
				SELECT @psFileName = fileUpload_fileName
				FROM ASRSysWorkflowInstanceValues
				WHERE instanceID = @piInstanceID
					AND elementID = @iElementID
					AND upper(ltrim(rtrim(isnull(identifier, '''')))) = @sItemIdentifier;
		
				SELECT fileUpload_file AS [file]
				FROM ASRSysWorkflowInstanceValues
				WHERE instanceID = @piInstanceID
					AND elementID = @iElementID
					AND upper(ltrim(rtrim(isnull(identifier, '''')))) = @sItemIdentifier;
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRWorkflowSubmitImmediatesAndGetSucceedingElements
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@succeedingElements	cursor varying output,
			@psTo				varchar(MAX)
		)
		AS
		BEGIN
			-- Action any immediate elements (Or, Decision and StoredData elements) and return the IDs of the workflow elements that 
			-- succeed them.
			-- This ignores connection elements.
			DECLARE
				@iTempID				integer,
				@iElementID				integer,
				@iElementType			integer,
				@iFlowCode				integer,
				@iTrueFlowType			integer,
				@iExprID				integer,
				@iResultType			integer,
				@sValue					varchar(MAX),
				@sResult				varchar(MAX),
				@fResult				bit,
				@dtResult				datetime,
				@fltResult				float,
				@iValue					integer,
				@iPrecedingElementType	integer, 
				@iPrecedingElementID	integer, 
				@iCount					integer,
				@iStepID				integer,
				@curRecipients			cursor,
				@sEmailAddress			varchar(MAX),
				@fDelegated				bit,
				@sDelegatedTo			varchar(MAX),
				@iSQLVersion			integer,
				@fStoredDataOK			bit, 
				@sStoredDataMsg			varchar(MAX), 
				@sStoredDataSQL			varchar(MAX), 
				@iStoredDataTableID		integer,
				@sStoredDataTableName	varchar(MAX),
				@iStoredDataAction		integer, 
				@iStoredDataRecordID	integer,
				@sStoredDataRecordDesc	varchar(MAX),
				@sStoredDataWebForms	varchar(MAX),
				@sStoredDataSaveForLater bit,
				@sSPName				varchar(MAX),
				@iNewRecordID			integer,
				@sEvalRecDesc			varchar(MAX),
				@iResult				integer,
				@iFailureFlows			integer,
				@fDeadlock				bit,
				@iErrorNumber			integer,
				@iRetryCount			integer,
				@iDEADLOCKERRORNUMBER	integer,
				@iMAXRETRIES			integer,
				@fIsDelegate			bit;
		
			SET @iDEADLOCKERRORNUMBER = 1205;
			SET @iMAXRETRIES = 5;
							
			SELECT @iSQLVersion = convert(float,substring(@@version,charindex(''-'',@@version)+2,2));
							
			DECLARE @elements table
			(
				elementID		integer,
				elementType		integer,
				processed		tinyint default 0,
				trueFlowType	integer,
				trueFlowExprID	integer
			);
							
			INSERT INTO @elements 
				(elementID,
				elementType,
				processed,
				trueFlowType,
				trueFlowExprID)
			SELECT SUCC.id,
				E.type,
				0,
				isnull(E.trueFlowType, 0),
				isnull(E.trueFlowExprID, 0)
			FROM [dbo].[udfASRGetSucceedingWorkflowElements](@piElementID, 0) SUCC
			INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID;
				
			SELECT @iCount = COUNT(*)
			FROM @elements
			WHERE (elementType = 4 OR (@iSQLVersion >= 9 AND elementType = 5) OR elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
				AND processed = 0;
		
			WHILE @iCount > 0
			BEGIN
				UPDATE @elements
				SET processed = 1
				WHERE processed = 0;
		
				-- Action any succeeding immediate elements (Decision, Or and StoredData elements)
				DECLARE immediateCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.elementID,
					E.elementType,
					E.trueFlowType, 
					E.trueFlowExprID
				FROM @elements E
				WHERE (E.elementType = 4 OR (@iSQLVersion >= 9 AND E.elementType = 5) OR E.elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND E.processed = 1;
		
				OPEN immediateCursor;
				FETCH NEXT FROM immediateCursor INTO 
					@iElementID, 
					@iElementType, 
					@iTrueFlowType, 
					@iExprID;
				WHILE (@@fetch_status = 0)
				BEGIN
					-- Submit the immediate elements, and get their succeeding elements
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 3,
						ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
						ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
						ASRSysWorkflowInstanceSteps.message = '''',
						ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
		
					SET @iFlowCode = 0;
		
					IF @iElementType = 4 -- Decision
					BEGIN
						IF @iTrueFlowType = 1
						BEGIN
							-- Decision Element flow determined by a calculation
							EXEC [dbo].[spASRSysWorkflowCalculation]
								@piInstanceID,
								@iExprID,
								@iResultType OUTPUT,
								@sResult OUTPUT,
								@fResult OUTPUT,
								@dtResult OUTPUT,
								@fltResult OUTPUT, 
								0;
		
							SET @iValue = convert(integer, @fResult);
						END
						ELSE
						BEGIN
							-- Decision Element flow determined by a button in a preceding web form
							SET @iPrecedingElementType = 4; -- Decision element
							SET @iPrecedingElementID = @iElementID;
		
							WHILE (@iPrecedingElementType = 4)
							BEGIN
								SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
									@iPrecedingElementType = isnull(WE.type, 0)
								FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPrecedingElementID) PE
								INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
								INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
									AND WIS.instanceID = @piInstanceID;
		
								SET @iPrecedingElementID = @iTempID;
							END
							
							SELECT @sValue = ISNULL(IV.value, ''0'')
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
							WHERE IV.elementID = @iPrecedingElementID
							AND IV.instanceid = @piInstanceID
								AND E.ID = @iElementID;
		
							SET @iValue = 
								CASE
									WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
									ELSE 0
								END;
						END
						
						IF @iValue IS null SET @iValue = 0;
						SET @iFlowCode = @iValue;
		
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
					END
					ELSE IF @iElementType = 7 -- Or
					BEGIN
						EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @piInstanceID, @iElementID;
					END
					ELSE IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
					BEGIN
						SET @fStoredDataOK = 1;
						SET @sStoredDataMsg = '''';
						SET @sStoredDataRecordDesc = '''';
		
						EXEC [spASRGetStoredDataActionDetails]
							@piInstanceID,
							@iElementID,
							@sStoredDataSQL			OUTPUT, 
							@iStoredDataTableID		OUTPUT,
							@sStoredDataTableName	OUTPUT,
							@iStoredDataAction		OUTPUT, 
							@iStoredDataRecordID	OUTPUT;
		
						IF @iStoredDataAction = 0 -- Insert
						BEGIN
							SET @sSPName  = ''sp_ASRInsertNewRecord_'' + convert(varchar(255), @iStoredDataTableID);
		
							SET @iRetryCount = 0;
							SET @fDeadlock = 1;
		
							WHILE @fDeadlock = 1
							BEGIN
								SET @fDeadlock = 0;
								SET @iErrorNumber = 0;
		
								BEGIN TRY
									EXEC @sSPName
										@iNewRecordID  OUTPUT, 
										@sStoredDataSQL;
		
									SET @iStoredDataRecordID = @iNewRecordID;
								END TRY
								BEGIN CATCH
									SET @iErrorNumber = ERROR_NUMBER();
		
									IF @iErrorNumber = @iDEADLOCKERRORNUMBER
									BEGIN
										IF @iRetryCount < @iMAXRETRIES
										BEGIN
											SET @iRetryCount = @iRetryCount + 1;
											SET @fDeadlock = 1;
											--Sleep for 5 seconds
											WAITFOR DELAY ''00:00:05'';
										END
										ELSE
										BEGIN
											SET @fStoredDataOK = 0;
											SET @sStoredDataMsg = ERROR_MESSAGE();
										END
									END
									ELSE
									BEGIN
										SET @fStoredDataOK = 0;
										SET @sStoredDataMsg = ERROR_MESSAGE();
									END
								END CATCH
							END
						END
						ELSE IF @iStoredDataAction = 1 -- Update
						BEGIN
							SET @sSPName  = ''sp_ASRUpdateRecord_'' + convert(varchar(255), @iStoredDataTableID);
		
							SET @iRetryCount = 0;
							SET @fDeadlock = 1;
		
							WHILE @fDeadlock = 1
							BEGIN
								SET @fDeadlock = 0;
								SET @iErrorNumber = 0;
		
								BEGIN TRY
									EXEC @sSPName
										@iResult OUTPUT,
										@sStoredDataSQL,
										@sStoredDataTableName,
										@iStoredDataRecordID,
										null;
								END TRY
								BEGIN CATCH
									SET @iErrorNumber = ERROR_NUMBER();
		
									IF @iErrorNumber = @iDEADLOCKERRORNUMBER
									BEGIN
										IF @iRetryCount < @iMAXRETRIES
										BEGIN
											SET @iRetryCount = @iRetryCount + 1;
											SET @fDeadlock = 1;
											--Sleep for 5 seconds
											WAITFOR DELAY ''00:00:05'';
										END
										ELSE
										BEGIN
											SET @fStoredDataOK = 0;
											SET @sStoredDataMsg = ERROR_MESSAGE();
										END
									END
									ELSE
									BEGIN
										SET @fStoredDataOK = 0;
										SET @sStoredDataMsg = ERROR_MESSAGE();
									END
								END CATCH
							END
						END
						ELSE IF @iStoredDataAction = 2 -- Delete
						BEGIN
							EXEC spASRRecordDescription
								@iStoredDataTableID,
								@iStoredDataRecordID,
								@sStoredDataRecordDesc OUTPUT;
		
							SET @sSPName  = ''sp_ASRDeleteRecord_'' + convert(varchar(255), @iStoredDataTableID);
		
							SET @iRetryCount = 0;
							SET @fDeadlock = 1;
		
							WHILE @fDeadlock = 1
							BEGIN
								SET @fDeadlock = 0;
								SET @iErrorNumber = 0;
		
								BEGIN TRY
									EXEC @sSPName
										@iResult OUTPUT,
										@sStoredDataTableName,
										@iStoredDataRecordID;
								END TRY
								BEGIN CATCH
									SET @iErrorNumber = ERROR_NUMBER();
		
									IF @iErrorNumber = @iDEADLOCKERRORNUMBER
									BEGIN
										IF @iRetryCount < @iMAXRETRIES
										BEGIN
											SET @iRetryCount = @iRetryCount + 1;
											SET @fDeadlock = 1;
											--Sleep for 5 seconds
											WAITFOR DELAY ''00:00:05'';
										END
										ELSE
										BEGIN
											SET @fStoredDataOK = 0;
											SET @sStoredDataMsg = ERROR_MESSAGE();
										END
									END
									ELSE
									BEGIN
										SET @fStoredDataOK = 0;
										SET @sStoredDataMsg = ERROR_MESSAGE();
									END
								END CATCH
							END
						END
						ELSE
						BEGIN
							SET @fStoredDataOK = 0;
							SET @sStoredDataMsg = ''Unrecognised data action.'';
						END
		
						IF (@fStoredDataOK = 1)
							AND ((@iStoredDataAction = 0)
								OR (@iStoredDataAction = 1))
						BEGIN
		
							exec [dbo].[spASRStoredDataFileActions]
								@piInstanceID,
								@iElementID,
								@iStoredDataRecordID;
						END
		
						IF @fStoredDataOK = 1
						BEGIN
							SET @sStoredDataMsg = ''Successfully '' +
								CASE
									WHEN @iStoredDataAction = 0 THEN ''inserted''
									WHEN @iStoredDataAction = 1 THEN ''updated''
									ELSE ''deleted''
								END + '' record'';
		
							IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
							BEGIN
								IF @iStoredDataRecordID > 0 
								BEGIN	
									EXEC [dbo].[spASRRecordDescription] 
										@iStoredDataTableID,
										@iStoredDataRecordID,
										@sEvalRecDesc OUTPUT
									IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc;
								END
							END
		
							IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')'';
		
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = convert(varchar(255), @iStoredDataRecordID), 
								ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @iElementID
								AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
								AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0;
		
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 3,
								ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.message = @sStoredDataMsg
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;
						END
						ELSE
						BEGIN
							-- Check if the failed element has an outbound flow for failures.
							SELECT @iFailureFlows = COUNT(*)
							FROM ASRSysWorkflowElements Es
							INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
								AND Ls.startOutboundFlowCode = 1
							WHERE Es.ID = @iElementID
								AND Es.type = 5; -- 5 = StoredData
		
							IF @iFailureFlows = 0
							BEGIN
								UPDATE ASRSysWorkflowInstanceSteps
								SET status = 4,	-- 4 = failed
									message = @sStoredDataMsg,
									failedCount = isnull(failedCount, 0) + 1,
									completionCount = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID;
		
								UPDATE ASRSysWorkflowInstances
								SET status = 2	-- 2 = error
								WHERE ID = @piInstanceID;
							END
							ELSE
							BEGIN
								UPDATE ASRSysWorkflowInstanceSteps
								SET status = 8,	-- 8 = failed action
									message = @sStoredDataMsg,
									failedCount = isnull(failedCount, 0) + 1,
									completionCount = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID;
		
								INSERT INTO @elements 
									(elementID,
									elementType,
									processed,
									trueFlowType,
									trueFlowExprID)
								SELECT SUCC.id,
									E.type,
									0,
									isnull(E.trueFlowType, 0),
									isnull(E.trueFlowExprID, 0)
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 1) SUCC
								INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
								WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements);
							END
						END
					END
		
					IF (@iElementType <> 5) OR (@fStoredDataOK = 1)
					BEGIN
						-- Get this immediate element''s succeeding elements
						INSERT INTO @elements 
							(elementID,
							elementType,
							processed,
							trueFlowType,
							trueFlowExprID)
						SELECT SUCC.id,
							E.type,
							0,
							isnull(E.trueFlowType, 0),
							isnull(E.trueFlowExprID, 0)
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, @iFlowCode) SUCC
						INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
						WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements);
					END
		
					FETCH NEXT FROM immediateCursor INTO 
						@iElementID, 
						@iElementType, 
						@iTrueFlowType, 
						@iExprID;
				END
				CLOSE immediateCursor;
				DEALLOCATE immediateCursor;
		
				UPDATE @elements
				SET processed = 2
				WHERE processed = 1;
		
				SELECT @iCount = COUNT(*)
				FROM @elements
				WHERE (elementType = 4 OR (@iSQLVersion >= 9 AND elementType = 5) OR elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND processed = 0;
			END
		
			SELECT @iCount = COUNT(*)
			FROM @elements
			WHERE elementType = 2; -- 2=WebForm
		
			IF (@iCount > 0) AND len(ltrim(rtrim(@psTo))) > 0 
			BEGIN
				SELECT @iStepID = ASRSysWorkflowInstanceSteps.ID
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID;
		
				DECLARE @recipients TABLE (
					emailAddress	varchar(MAX),
					delegated		bit,
					delegatedTo		varchar(MAX),
					isDelegate		bit
				);
		
				exec [dbo].[spASRGetWorkflowDelegates] 
					@psTo, 
					@iStepID, 
					@curRecipients output;
				FETCH NEXT FROM @curRecipients INTO 
						@sEmailAddress,
						@fDelegated,
						@sDelegatedTo,
						@fIsDelegate;
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO @recipients
						(emailAddress,
						delegated,
						delegatedTo,
						isDelegate)
					VALUES (
						@sEmailAddress,
						@fDelegated,
						@sDelegatedTo,
						@fIsDelegate
					);
					
					FETCH NEXT FROM @curRecipients INTO 
							@sEmailAddress,
							@fDelegated,
							@sDelegatedTo,
							@fIsDelegate;
				END
				CLOSE @curRecipients;
				DEALLOCATE @curRecipients;
		
				DELETE FROM ASRSysWorkflowStepDelegation
				WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
					FROM ASRSysWorkflowInstanceSteps
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID IN 
							(SELECT E.elementID
							FROM @elements E
							WHERE E.elementType = 2) -- 2 = WebForm
						AND (ASRSysWorkflowInstanceSteps.status = 0
							OR ASRSysWorkflowInstanceSteps.status = 2
							OR ASRSysWorkflowInstanceSteps.status = 6
							OR ASRSysWorkflowInstanceSteps.status = 8
							OR ASRSysWorkflowInstanceSteps.status = 3));
		
				INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
				SELECT DISTINCT RECS.emailAddress, WIS.ID
				FROM @recipients RECS, 
					ASRSysWorkflowInstanceSteps WIS
				WHERE RECS.isDelegate = 1
					AND WIS.instanceID = @piInstanceID
						AND WIS.elementID IN 
							(SELECT E.elementID
							FROM @elements E
							WHERE E.elementType = 2) -- 2 = WebForm
						AND (WIS.status = 0
							OR WIS.status = 2
							OR WIS.status = 6
							OR WIS.status = 8
							OR WIS.status = 3);
			END
		
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 1,
				ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
				ASRSysWorkflowInstanceSteps.completionDateTime = null,
				ASRSysWorkflowInstanceSteps.userEmail = CASE
					WHEN (SELECT ASRSysWorkflowElements.type 
						FROM ASRSysWorkflowElements 
						WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @psTo -- 2 = Web Form element
					ELSE ASRSysWorkflowInstanceSteps.userEmail
				END
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID IN 
					(SELECT E.elementID
					FROM @elements E
					WHERE E.elementType <> 7 -- 7 = Or
						AND (E.elementType <> 5 OR @iSQLVersion <= 8) -- 5 = StoredData
						AND E.elementType <> 4) -- 4 = Decision
				AND (ASRSysWorkflowInstanceSteps.status = 0
					OR ASRSysWorkflowInstanceSteps.status = 2
					OR ASRSysWorkflowInstanceSteps.status = 6
					OR ASRSysWorkflowInstanceSteps.status = 8
					OR ASRSysWorkflowInstanceSteps.status = 3);
		
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 2
			WHERE ASRSysWorkflowInstanceSteps.id IN (
				SELECT ASRSysWorkflowInstanceSteps.ID
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND ASRSysWorkflowElements.type = 2);
		
			-- Return the cursor of succeeding elements. 
			SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
				SELECT elementID 
				FROM @elements E
				WHERE E.elementType <> 7 -- 7 = Or
					AND E.elementType <> 4 -- 4 = Decision
					AND (E.elementType <> 5 OR @iSQLVersion <= 8); -- 5 = StoredData
		
			OPEN @succeedingElements;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRWorkflowTriggering
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowTriggering]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowTriggering];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRWorkflowTriggering]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRWorkflowTriggering]
		(
			@pfTrigger bit OUTPUT
		)
		AS
		BEGIN
			DECLARE @sInProgress varchar(MAX);
		
			SET @pfTrigger = 0;
		
			SELECT @sInProgress = isnull(settingValue, ''0'')
			FROM ASRSysSystemSettings
			WHERE section = ''workflow''
				AND settingKey = ''triggering'';
		
			IF @sInProgress = ''0''
			BEGIN
				SET @pfTrigger = 1;
		
				DELETE FROM ASRSysSystemSettings
				WHERE section = ''workflow''
					AND settingKey = ''triggering'';
		
				INSERT INTO ASRSysSystemSettings
					(section, settingKey, settingValue)
				VALUES 
					(''workflow'', ''triggering'', ''1'');
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRWorkflowValidRecord
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowValidRecord]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowValidRecord];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRWorkflowValidRecord]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'Alter PROCEDURE [dbo].[spASRWorkflowValidRecord]
			@piInstanceID				integer,
			@piRecordType				integer,
			@piRecordID					integer,
			@sElementIdentifier			varchar(MAX),
			@sElementItemIdentifier		varchar(MAX),
			@pfValid					bit		OUTPUT
		AS
		BEGIN
			DECLARE
				@iTableID				integer,
				@iWorkflowID			integer,
				@iElementType			integer;
		
			SET @pfValid = 0;
		
			SELECT @iWorkflowID = WF.ID,
				@iTableID = 
					CASE
						WHEN @piRecordType = 4 THEN isnull(WF.baseTable, 0)
						ELSE 0
					END
			FROM ASRSysWorkflows WF
			INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
				AND WFI.ID = @piInstanceID;
		
			IF @piRecordType = 0
			BEGIN
				-- Initiator''s record
				SELECT @iTableID = convert(integer, ISNULL(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel'';
		
				IF @iTableID = 0
				BEGIN
					SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_WORKFLOW''
					AND parameterKey = ''Param_TablePersonnel'';
				END
			END
		
			IF @piRecordType = 1
			BEGIN
				-- Identified record
				SELECT @iElementType = ASRSysWorkflowElements.type,
					@iTableID = 
						CASE
							WHEN ASRSysWorkflowElements.type = 5 THEN isnull(ASRSysWorkflowElements.dataTableID, 0)
							ELSE 0
						END
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sElementIdentifier)));
		
				IF @iElementType = 2
				BEGIN
					 -- WebForm
					SELECT @iTableID = WFEI.tableID
					FROM ASRSysWorkflowElementItems WFEI
					INNER JOIN ASRSysWorkflowElements WFE ON WFEI.elementID = WFE.ID
						AND WFE.identifier = @sElementIdentifier
						AND WFE.workflowID = @iWorkflowID
					WHERE WFEI.identifier = @sElementItemIdentifier;
				END
			END
		
			EXEC [dbo].[spASRWorkflowValidTableRecord]
				@iTableID,
				@piRecordID,
				@pfValid	OUTPUT;
		END';

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRWorkflowValidTableRecord
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowValidTableRecord]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowValidTableRecord];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRWorkflowValidTableRecord]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRWorkflowValidTableRecord]
			@piTableID	integer,
			@piRecordID	integer,
			@pfValid	bit			OUTPUT
		AS
		BEGIN
			DECLARE	@sSQL	nvarchar(MAX),
					@sParam	nvarchar(500);
				
			SET @pfValid = 0;
		
			IF EXISTS (SELECT *
				FROM dbo.sysobjects
				WHERE id = object_id(N''[dbo].[udf_ASRWorkflowValidTableRecord]'')
					AND OBJECTPROPERTY(id, N''IsScalarFunction'') = 1)
			BEGIN
				SET @sSQL = ''SET @pfValid = [dbo].[udf_ASRWorkflowValidTableRecord]('' 
					+ convert(nvarchar(100), @piTableID) 
					+ '', '' 
					+ convert(nvarchar(100), @piRecordID)
					+ '')'';
				SET @sParam = N''@pfValid bit OUTPUT'';
				EXEC sp_executesql @sSQL, @sParam, @pfValid OUTPUT;
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;



/* ------------------------------------------------------------- */
PRINT 'Step 11 - New Shared Table Transfer Types'

	-- SMP
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 6 AND TransferFieldID = 20
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (20,6,0,''MATB1 Received Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (21,6,0,''Actual Date of Birth'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- SPP
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 7 AND TransferFieldID = 22
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (22,7,0,''SC3 Received Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (23,7,0,''Still Birth'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- SAP
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 8 AND TransferFieldID = 19
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (19,8,0,''Matching Certificate Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (20,8,0,''Child Expected Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (21,8,0,''Actual Placed Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END	

	-- Generic Absence
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 72
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (72, ''Absence'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,72,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,72,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,72,1,''Absence Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,72,0,''Absence Reason'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,72,1,''Start Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,72,0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,72,1,''Start Session'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,72,0,''End Session'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,72,0,''Hours Duration'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,72,0,''Start Time'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,72,0,''End Time'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Adjustment for v0.8 of Fulcrum absence spec
	IF NOT EXISTS (SELECT TransferFieldID
			FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 72 AND TransferFieldID = 11)
	BEGIN	
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,72,1,''Absence ID'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Some test databases have had previous script run on this.
	SELECT @NVarCommand = 'UPDATE [ASRSysAccordTransferFieldDefinitions]
			SET [AlwaysTransfer] = 1
			WHERE [TransferTypeID] = 72 AND [TransferFieldID] = 2;'
	EXEC sp_executesql @NVarCommand


	-- SAP Adoption
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 73
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (73, ''SPP Adoption'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,73,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,73,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,73,1,''SC4 Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,73,0,''Child Expected Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,73,0,''Actual Placed Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,73,0,''Start SSP Leave'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END


	-- Working Patterns
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 74
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (74, ''Working Pattern'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,74,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,74,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,74,1,''Effective Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,74,1,''Working Pattern'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Adjustment for v5 of Fulcrum absence spec
	SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET Description = ''Working Pattern AM'' WHERE TransferTypeID = 74 AND TransferFieldID = 3'
	EXEC sp_executesql @NVarCommand

	IF NOT EXISTS (SELECT TransferFieldID
			FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 74 AND TransferFieldID = 4)
	BEGIN	
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,74,1,''Working Pattern PM'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Keeping in Touch Days (Maternity)
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 75
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (75, ''Keeping in Touch Days (Maternity)'' ,0,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,75,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,75,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,75,1,''MATB1 Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,75,1,''Start Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,75,1,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,75,1,''Reason'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Keeping in Touch Days (Adoption)
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 76
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (76, ''Keeping in Touch Days (Adoption)'' ,0,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,76,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,76,1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,76,1,''Child Expected Date Adoption'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,76,1,''Start Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,76,1,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,76,1,''Reason'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Adjustment for v0.9 of Fulcrum absence spec
	IF NOT EXISTS (SELECT TransferFieldID
			FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 8 AND TransferFieldID = 22)
	BEGIN	
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (22,8,0,''Work up to Placement'',0,0,2,0,0);'
		EXEC sp_executesql @NVarCommand
	END

	IF NOT EXISTS (SELECT TransferFieldID
			FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 73 AND TransferFieldID = 6)
	BEGIN	
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions 
		 (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,73,0,''Work up to Placement'',0,0,2,0,0);'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 12 - Adding New Standard Colours'

DECLARE @iMaxColOrder integer

SELECT @iMaxColOrder = MAX(ColOrder) from ASRSysColours
IF @iMaxColOrder IS NULL SET @iMaxColOrder = 1

IF NOT EXISTS (SELECT * FROM ASRSysColours WHERE ColDesc = 'Midnight Blue')
BEGIN
	SET @iMaxColOrder = @iMaxColOrder + 1
	SELECT @NVarCommand = 'INSERT INTO ASRSYSCOLOURS (ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
			VALUES (' + convert(varchar(10), @iMaxColOrder) + ', 6697779, ''Midnight Blue'', 3, 0)'
	EXEC sp_executesql @NVarCommand
END

IF NOT EXISTS (SELECT * FROM ASRSysColours WHERE ColDesc = 'Dolphin Blue')
BEGIN
	SET @iMaxColOrder = @iMaxColOrder + 1	
	SELECT @NVarCommand = 'INSERT INTO ASRSYSCOLOURS (ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
				VALUES (' + convert(varchar(10), @iMaxColOrder) + ', 16248553, ''Dolphin Blue'', 3, 0)'
	EXEC sp_executesql @NVarCommand
END
	
IF NOT EXISTS (SELECT * FROM ASRSysColours WHERE ColDesc = 'Pale Grey')
BEGIN
	SET @iMaxColOrder = @iMaxColOrder + 1	
	SELECT @NVarCommand = 'INSERT INTO ASRSYSCOLOURS (ColOrder, ColValue, ColDesc, WordColourIndex, CalendarLegendColour)
				VALUES (' + convert(varchar(10), @iMaxColOrder + 1) + ', 15988214, ''Pale Grey'', 15, 0)'
	EXEC sp_executesql @NVarCommand
END	


/* ------------------------------------------------------------- */
PRINT 'Step 13 - Expression amendments'

	SELECT @NVarCommand = 'UPDATE [dbo].[ASRSysFunctions] SET [spName] = ''sp_ASRFn_ConvertToProperCase'' WHERE [FunctionID] = 12;'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 14 - Updating System Permissions Icons'

/* Updating System Permissions Icon for Module Access */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 1

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 1

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000000000000000000000000000000000000000000000000070306000F060D000E0D0E00100F0F001411130024162000281523002A1B26003B162D003D172F002C2C2C0032282E00302F2F00372131003B3B3A00401F3200491A360061224E007329540072396100404040004F4F4E005648510051505000575655005E5E5E007C496F00606060006766650077767600787776007E737A0081376D008F467A008D507C00964F82009A5586009F5E8D0090648300A8759900AD759D00AB7F9900448BF000468CF000488DF0004F92F0005093F1005295F1005798F2005E9CF30070A7F30071A9F50074ABF50076ACF500918F8E00979796009D9C9A009E9E9E00A9829C00A09F9D00B684A700B68DA500B786A900B887A900BB8CAD00B799AD00A2A2A100A8A6A500A8A8A800B3B2B100C195B300C297B400C399B600C59BB700C7A2BB00C8A1BC00CAA5BE00B6B3D90080AFF20080B1F5008BB9F70092BEF800D5B5C900D4B8CA00D7B9CC00B0D1FA00D1D0CF00D4D3D300D7D6D500D9D8D800DFDEDE00E1D2DC00E0DFDE00E5E4E400EFE1E900EBEAEA00ECEBEB00F5F5F500FBFAFA00FDFDFD004CB0000059CF000067F0000078FF11008AFF31009CFF5100AEFF7100C0FF9100D2FFB100E4FFD100FFFFFF0000000000262F0000405000005A700000749000008EB00000A9CF0000C2F00000D1FF1100D8FF3100DEFF5100E3FF7100E9FF9100EFFFB100F6FFD100FFFFFF00000000002F26000050410000705B000090740000B08E0000CFA90000F0C30000FFD21100FFD83100FFDD5100FFE47100FFEA9100FFF0B100FFF6D100FFFFFF00000000002F1400005022000070300000903E0000B04D0000CF5B0000F0690000FF791100FF8A3100FF9D5100FFAF7100FFC19100FFD2B100FFE5D100FFFFFF00000000002F030000500400007006000090090000B00A0000CF0C0000F00E0000FF201200FF3E3100FF5C5100FF7A7100FF979100FFB6B100FFD4D100FFFFFF00000000002F000E00500017007000210090002B00B0003600CF004000F0004900FF115A00FF317000FF518600FF719C00FF91B200FFB1C800FFD1DF00FFFFFF00000000002F0020005000360070004C0090006200B0007800CF008E00F000A400FF11B300FF31BE00FF51C700FF71D100FF91DC00FFB1E500FFD1F000FFFFFF00000000002C002F004B0050006900700087009000A500B000C400CF00E100F000F011FF00F231FF00F451FF00F671FF00F791FF00F9B1FF00FBD1FF00FFFFFF00000000001B002F002D0050003F007000520090006300B0007600CF008800F0009911FF00A631FF00B451FF00C271FF00CF91FF00DCB1FF00EBD1FF00FFFFFF000000000008002F000E005000150070001B0090002100B0002600CF002C00F0003E11FF005831FF007151FF008C71FF00A691FF00BFB1FF00DAD1FF00FFFFFF0000001800040400083B4A260000000000001C0F3944371F0D0C4A26000000000000154360615A3C1F004D400000000000001545646462571D022122000000000000001A6F6F635D160E214B4F332D50000000013A5A5846051B2800312F2D360000000A185B591907230000302D2B51000000091E64633806545F002B2B2B52000000110B5A5E15204C41554E0000000000001310030017003F254953000000000000132A12140000292448530056000000003E42275C0000003D47530035353400000000000000000000000000322C2B000000000000000000000000002B2B2B000000000000000000000000002E2B2B00000000000000000000000000000000C01F0000801F0000801F0000801F0000C0010000C0210000C0610000C0210000C00F0000C10F0000C30B0000C3880000FFF80000FFF80000FFF80000FFFF000000
END

/* Updating System Permissions Icon for Batch Jobs */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 2

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 2

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000009C1852BD6B6363636B6B6B6B6B7373737373737B7B7B7B847B5A848484848C8C8C845A8C8C8C94949494949C949C9C9C9C9C9C9CA59CA5A59CADC6A5A5A5A5B5CEA5BDD6ADADADADC6DEADC6E7B5B5B5B5E7BDBD0808BDB573BDBDADBDBDBDC6AD4AC6C6C6CE0000CECECED6CEADD6D6D6DED6ADDEDEDEDEE7F7E7EFEFE7EFF7F7F7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100002C002C00000000100010000008A30059081C48B0208B1010122A5CA830C208051630489C481183080B0D1C54DC281123041520438A0C89C20184830C173A5C919085028E15267068C0F2240414387176303040810002235A9E1C4842808781140A30B039700502062C000CFC3080A9401302040200B041A0009B2B048258A0756B0016074E7EB840814283AC51B97A8510B669820F51076A18B0B6ADDF060328201DA0C1A0C0120904301890E0048B8000003B00
END

/* Updating System Permissions Icon for Calculations */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 21

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 21

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F7000018212121292929313129393931393131393931394231424231424A31426B39393939424239424A394A4A394A52424242424A52424A6B42525A4A4A6B4A5A5A4A5A63525A5A525A635263635A526B5A63635A6B735A6B84635263636363636B6B6373736B6B6B6B7373735263736373737373737B7B7B5A637B5A6B7B73847B7B7B7B8484847B8C848C948C8C8C8C94948C949C8C9CAD949494949C94949CAD9C9C9C9CA5A5A5A5A5A5ADADA5ADB5ADADADADB5B5ADB5BDB59494B5ADADB5B5B5B5B5BDBDBDBDBDBDC6C6B5BDC6C6C6C6CECECECECECED6D6D6D6D6DEDEDEE7E7E7F7F7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100004C002C00000000100010000008D1009908647204428B8108133201E1A0C10B8509774870C0E0C00E8802975480D0A0E28003216A10490863624503216E8878000080851A0425983CB000888C1B4482C8085180C906091C111858F1C3835115358C2CC02193E2010839659430EAC286070C409DCAF0D1638891203582682052B221020A44509C38418205901F0B9868748A234587BB1926F028A143E00E8A1F784498403802872005960CBC70C0840E223438244890C3450984461A1418B040848E20318C14509230090E110502002810C285058C0291D428B160C148810101003B00
END

/* Updating System Permissions Icon for Calendar Reports */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 24

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 24

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F7000063737B638C9C638CA56394AD6394B56B737B6B7B846B94B56B9CB56B9CBD6B9CC6737B8473848473848C738C8C7B848C7B8C9484949484949C8C949C8C9C9C8C9CA58CBDDE949CA594A5A594A5AD94C6DE94C6E794CEE79CA5AD9CCEE79CD6EFA5ADB5A5B5B5A5DEEFA5DEF7A5E7F7ADBDBDADE7F7ADEFFFB5C6CEB5CEE7B5EFFFBDC6C6BDC6CEBDCECEC6CECECED6DED6D6DED6DEDEDEDEE7DEE7E7E7E7E7E7E7EFE7EFEFEFEFEFEFEFF7EFF7F7F7F7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100003B002C00000000100010000008C00077081C48B0A0C18304532458980001820303220E1040114502132732663461628408111F3C6C08906023098F1F4272D8A0C18205920C1D428C485140809B00001430C0934183060F20088D2041020084057518DD2163C60C1A356AD8B87103878EAB130A8078F134EAD41B25325CD541C1C00A083E817EC521616C05033B5EAC60310303D51B3920E8C8A1E3ED8E193158BC8040D56A83BD3A2E2CD80195068B0933F0FE34C0170383822B42081C7B15C3CFA0438B4AA050A1C2050C190202003B00
END

/* Updating System Permissions Icon for Career Progression */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 39

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 39

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000009C3100FF31089C312929AD3939AD3939B54242B54A4AB54A4ABD5252BD5A5ABD6363BD6363C66B6BC67373C67B7BCE8484CE8C8CD69494D69C9CD6A5A5DEADADDEB5B5E7BDBDE7C6C6E7C6C6EFD6D6EFE7E7F7EFEFF7EFEFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100001E002C000000001000100000089B003D081CE86183860D04130E8CB0A0A1048504292898702182820A1005367060A1E38206193D2C5000214284040B424A30908081020302427A6840A0A6809B1939404870E066009C0937346800C167809F31090E8D60F428D281121F4C1080B46A52910E1E48B8C9F427CCA41A123C200A018203AA020668185821C150080F1E9CFD3900A3C00B071AC2958B760086882F0D082670530005810101003B00
END

/* Updating System Permissions Icon for CMG & Centrefile */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 20

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 20

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000009C3100FF316363636B6B6B7373738484849494949C9C9CADADADB5C6E7BDBDBDC6C6C6CECECED6DEEFE7EFF7EFEFEFF7F7F7F7FFFFFF4242FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000013002C000000001000100000088B00271C1848B0A0418109122A5CA8F080C00610234A6C20A181C30310326ADCA8F1A2C18F04053A1839A1A4C99307060858C9B2254B8102264078B00081819800724E887940C0839F407106C8C9530083050A0E1420C012C050003065D2CC49D5E9D39E401F580DC0B52BD1A349ABE6E49AB3A7D49A374B3AD56976265A0107D60220A8D2E58009544F0604003B00
END


/* Updating System Permissions Icon for Configuration */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 22

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 22

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700000000CE00FF0029947B31313131947B42BDB552BDA55ABDA563948C8C8C8C94949494949C9C9C8C9C9C9C9C9CA5A5A5A5A5BDD6AD9442ADADADADBDD6ADC6DEB59452B5A542C69439C6AD6BC6C6C6CE0000CEA539CEBD8CCECEBDCECEC6CECECED69418D69C18D69C21D69C29DE7B18DE9C31DEA529DEA531DEAD4ADEB552DEB55ADEB563DEEFF7E7AD31E7B531E7B552E7BD63E7EFF7EFB539EFBD39EFEFEFF7F7F7FF0000FF0800FFCE00FFEFADFFEFB5FFF7BDFFFFEFFFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100003E002C00000000100010000008A7007D081C48B0A0C18304333458C8B02143071F1240A040B1A2450A1A20285870B123C5090A1AC418399202C9912C0EACA8D0C00780811404DA9859A0808508224FC6F46183078A0D060410106983E4CE1B2946B4104100414EA33E6A90282103840A1A0B7D0480D943C508192254F4A89195E084154A455CF010352B8E811C468808EBA1C107B20D6A0CE4A1234709157A0562CD20E181610639746030CCD86EC1011D7A200C08003B00
END

/* Updating System Permissions Icon for Cross Tabs */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 3

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 3

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700002929295A5A5A63424A6B6B6B6B738C734A4A7B8CA5846339848494848CA5848CAD8494AD8C52108C7B088C94AD8C9CAD949CAD949CB59C94949C9C9C9CA5B59CADCEA5A5B5A5ADB5A5ADBDA5B5CEA5B5D6AD9C84ADA594ADAD84ADADB5ADADBDADB5BDADB5C6ADBDD6B58452B5B5BDB5B5C6B5BDC6B5BDD6B5BDDEB5C6DEBD6B10BDA510BDA584BDB584BDBDBDBDBDC6BDC6C6BDC6DEC6C6C6C6C6CECECEC6CECECECED6E7DEDEDEEFEFEFF79418FFDE10FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF2C00000000100010000008C20011448040F08183830B1224346080008514294EA010714284080D193254D8688042880F202E7CC0708124060B172840584001C48E973063BE34E1C082CB0D2C64ECE8D0A227871B3B5E3CB85062878A1C2C76ACD0C174848BA0102E90D8C140C5861D0D56683DF02228050C53658A7DF175AA0B174063A67D6101AC521D2D62E2008063078CB65397C685E922C084A018DCEA7D8963C08D1A2FEFBA155040C2CB0000D2DA0D1C56A664C01840C07821E3C566199F397BC0F0C186E9D3A851630808003B00
END

/* Updating System Permissions Icon for Custom Reports */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 12

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 12

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700004A00A55200A55200AD5A00AD6300AD6B08B57310B57318B57B21BD8421BD8429BD8C31BD8C39C68C42C69442C6944AC69C52C69C52CEA56BCEA56BD6AD6BD6AD73D6B57BD6B584D6B584DEBD84D6BD8CD6BD8CDEC694DEC69CDEC69CE7CE9CDECEA5E7CEADE7D6ADE7D6B5E7D6B5EFDEBDE7DEBDEFDEC6EFDECEEFE7CEEFE7D6F7E7DEF7EFDEF7EFE7F7F7E7F7F7EFF7F7EFFFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000032002C00000000100010000008AF003B74E040508341103164285CD861E1C20C273838542851218C172D2E7428116262C5182F0000081060C00002042A28D42003824B0013431448214106CB8B2D02C8289182458B170620A0B8C0D225049D0892229071008102051764BC609162808C9F0917AA5810D52801194AC3CA681075C589105F27CA48F8C0820C14233A14001B36694B9529466C30A016E4DD0A12021FA05BB7E5C4A50E2FB2880161A202B56B4F3476B88081830746330704003B00
END

/* Updating System Permissions Icon for Data Manager Intranet */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 19

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 19

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700000000CE00FF008C8C8C94949494949C949C9C9C9C9C9C9CA5A5A5A5A5BDD6ADADADADC6DEC6C6C6CE0000CECECEE7EFEFE7EFF7EFEFEFF7F7F7FF0000FFCE00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000015002C000000001000100000087A002B081C48B0A0C18304191858C8B021C3030E04245840B1A2C5050D120C2070B123458D06208814B960A4C907040C5400307081C00930275490B0D02404971564E69C5973C2489C3A65D20C6912684C9E2A03B43C38B4A04B07151C1880DA94C2D2A951A70E95401067C1080618284040B6AC59B258111E0C08003B00
END

/* Updating System Permissions Icon for Data Transfer */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 4

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 4

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000009C3100FF31089C31636363636B636B6B6B7373738484849494949C9C9CADADADB5C6E7BDBDBDC6C6C6CECECED6DEEFE7EFF7EFEFEFF7F7F7F7FFFFFF4242FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000015002C0000000010001000000892002B241848B0A041810B122A5CA83081C00710234A7C40E181C30412326ADCA8F1A2C18F0405421859A1644901284B262830A0A5CB96040404103040604D09111A284050B3824C01046C46184AB4E7CFA009063868C020C101032807FC944AB3024E9D087E6A9D398028519401C2868DBAB4E953AA625126B59A73674DB029D75E753B10A500822C5F0E28605280C90A0101003B00
END

/* Updating System Permissions Icon for Diary */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 5

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 5

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700003131316363639C9C9CC69400DEDEDEE7D69CE7E7E7EFEFEFF7F7F7FF0808FF1010FF1818FF2121FF2929FF3131FF3939FF4242FF4A4AFF5252FF5A5AFF6B6BFF7373FF7B7BFF9C9CFFADADFFC6C6FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100001A002C0000000010001000000890003568101040A04183040F263C88B06043860F351418407080C58B162B4EB44070828408101E3868C08060028B162A50F0085224C9050A4E0EB08021C3858F214732802933E54A9C2E77C6B4C81081D103480D28252A10805300080E0038A0D400538146A73AAD4AE02A82A748012825D0758041A35FC38E2D8BF5E9540362C9B2451B95EA5AB2032662DC8BB1404000003B00
END

/* Updating System Permissions Icon for Email Addresses */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 35

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 35

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700004A7BBD4A7BC6527BC65284C65A84C6638CCE639C6B63BD5263BD7B6B73B56B8CCE6B94CE6B9C736B9CAD6B9CD66BA5846BA5D66BC65A6BC67B7394CE739CCE739CD673A58473A5D673ADDE7B84BD7BA5D67BADDE7BB5DE7BC67384A5D684AD9C84ADD684ADDE84B5DE84B5E78CADDE8CB5DE8CB5E78CBDE78CCE7B94A5D694B5DE94BDE794CE8C94D68494D68C94D6949C9CCE9CBDDE9CBDE79CC6E79CD6949CDE9CA5ADD6A5B5D6A5BDDEA5BDE7A5C6E7A5D694A5DEA5A5DEB5ADB5D6ADB5DEADBDDEADC6E7ADCEE7ADCEEFADD69CADDEA5ADDEADB5BDDEB5C6DEB5CEE7B5CEEFB5D6EFB5E7B5BDC6E7BDD6EFBDD6F7BDDEF7BDE7BDC6CEDEC6CEE7C6D6EFCED6E7CED6EFCEDEEFCEE7F7CEEFCED6DEEFD6E7EFD6E7F7D6EFDEDEE7F7DEEFF7DEF7DEE7EFF7E7F7FFEFF7EFEFF7FFF7F7FFF7FFF7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000068002C00000000100010000008EE00D1081C88A60C95203892683943D04C16234CC604A1A022860A105308767951E3450F023196385912424399814678B020820000960D18560859106620131A11504800E0040204113A289C14D885C5810E3C406C10B1A144081C03C718C1C98205080D2744380D2250CA8707160C306820C0C3840922565CD0820686173170E38AF1D2448109195C8F4CB112A66F5F2B53C214D8704285972361AA34F15BA54A5F023F555039E2C58B9623476C4CA9EC858003C24E7C6CD98264CA962948466FF1BC34888FCB55B4C8AEF25ACBE0132192DCB831E58AEFDF37AE4C28B12288970436922B579E2105152D64D00404003B00
END

/* Updating System Permissions Icon for Email Groups */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 36

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 36

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700004A7BBD4A7BC6527BC65284C65A84C6638CCE639C6B63BD5263BD7B6B73B56B8CCE6B94CE6B9C736B9CAD6B9CD66BA5846BA5D66BC65A6BC67B7394CE739CCE739CD673A58473A5D673ADDE7B84BD7BA5D67BADDE7BB5DE7BC67384A5D684AD9C84ADD684ADDE84B5DE84B5E78CADDE8CB5DE8CB5E78CBDE78CCE7B94A5D694B5DE94BDE794CE8C94D68494D68C94D6949C9CCE9CBDDE9CBDE79CC6E79CD6949CDE9CA5ADD6A5B5D6A5BDDEA5BDE7A5C6E7A5D694A5DEA5A5DEB5ADB5D6ADB5DEADBDDEADC6E7ADCEE7ADCEEFADD69CADDEA5ADDEADB5BDDEB5C6DEB5CEE7B5CEEFB5D6EFB5E7B5BDC6E7BDD6EFBDD6F7BDDEF7BDE7BDC6CEDEC6CEE7C6D6EFCED6E7CED6EFCEDEEFCEE7F7CEEFCED6DEEFD6E7EFD6E7F7D6EFDEDEE7F7DEEFF7DEF7DEE7EFF7E7F7FFEFF7EFEFF7FFF7F7FFF7FFF7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000068002C00000000100010000008EE00D1081C88A60C95203892683943D04C16234CC604A1A022860A105308767951E3450F023196385912424399814678B020820000960D18560859106620131A11504800E0040204113A289C14D885C5810E3C406C10B1A144081C03C718C1C98205080D2744380D2250CA8707160C306820C0C3840922565CD0820686173170E38AF1D2448109195C8F4CB112A66F5F2B53C214D8704285972361AA34F15BA54A5F023F555039E2C58B9623476C4CA9EC858003C24E7C6CD98264CA962948466FF1BC34888FCB55B4C8AEF25ACBE0132192DCB831E58AEFDF37AE4C28B12288970436922B579E2105152D64D00404003B00
END

/* Updating System Permissions Icon for Email Queue */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 18

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 18

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700000084CE008CD60094DE009CDE00A5E700ADE700B5EF108CCE108CD610ADE710B5EF10BDEF1894D6189CDE18BDEF218CBD297B9C29B5E729B5EF31ADE731C6F739BDEF42A5D64AA5CE52ADDE52CEF75AC6EF5ACEF7636B6B6394AD63B5E763CEF76B6B6B6B73736B7B8C6BA5BD6BBDE76BBDEF6BD6F773737373ADC67B7B7B7B8484848484848C8C84C6E78C8C8C94948C94949494D6F79C9C949C9C9CA5A5A5ADADADADDEF7B5B5B5BDBDBDBDE7F7BDEFFFC6C6C6C6DEE7CECECED6D6D6D6EFF7DEEFF7E7E7E7EFEFEFEFFFFFF7F7F7F7F7FFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000047002C00000000100010000008C2008F081C48B0A0C183046FDCA85183068D193360C078E162C50A15296EEC308250A0911D1969D4E078D048C31335823414625088C820288304B90183E54021306EC80C5143884F1C2F7C08F4F102874F2120681C15B263850F1F233464C810034852225889ECE8E102450503261C28F8C08146D61D4188F0A850A0C0100D05127498418469D61C090844C81161C0800B3382F42018C3AF001B130408781011868B8A2B440808C020078900012020F4008001090C000E0CE9D8E2000000167E1C0908003B00
END

/* Updating System Permissions Icon for Envelope & Label Templates */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 30

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 30

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700005A6B848484848C8C8C9494949C9C9CA5A5A5AD1818ADADADB51818B54242B54A29B57B7BB5B5B5BD2121BDBDBDC64A4AC68C39C6BDA5C6C6C6CE7329CE9C39CECECED66B29D69431D69C9CD6AD39D6D6D6DE7B29DEDEDEE7AD29E7CECEE7EFEFEFB529EFB531EFBD42EFC642EFC64AEFCE7BEFEFEFF7D65AF7D663F7E79CF7E7BDF7EFDEF7F7F7FFE7A5FFEFB5FFEFBDFFF7DEFFF7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000032002C00000000100010000008AD0065081C48B0A0C1830409282CC0B0E181870F19302020810542812C244C2C70C0E24116070A4CE410F24341160F1A2CE0C0A0C0870F0706981C68A1814D0C2D5F7E7020A082C00E1136D834D0D2845113120254E88022C58A09061238287014E9D2A62248A850E0612A0B16483F30751142C409085FBD7EA8C8F4458811282048F8C0622A870A2C3A7480118244DCAF15384C85D8216F89B8121D28764030C3071314205C1408804264840101003B00
END

/* Updating System Permissions Icon for Envelopes & Labels */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 29

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 29

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F7000084319C9C63AD9C63B5A563B5A56BB5ADADADB58CC6BD8CC6BD94C6BD94CEBD9CCEBDBDBDC6ADCECEADD6CEB5D6CECECED6D6D6DEDEDEE7E7E7E7E7EFE7EFEFEFEFEFEFF7F7F7F7F7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000019002C00000000100010000008700033081C48B0A0C18308132A4CB8A0A1C3871017149040B1A2458B050A285480A0C1858C0931000010A0024884174696046980804B02030618B8800001030A270B5ED879A14205092C5FC6F4D8F367849C0279169510E1A8C6811878FAA4180102529E167E3285F020A3D7AF60BD0604003B00
END

/* Updating System Permissions Icon for Event Log */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 17

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 17

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000007BC60084CE008CD60094DE009CDE00A5E700ADE700B5EF108CD610ADE710B5EF10BDEF1894D6189CDE18BDEF296B9429B5E729B5EF31ADE731C6F7399CD639BDEF4AA5D64AB5E752525252ADDE52CEF75A5A5A5AC6EF6363636B6B6B6BBDE76BBDEF6BCEF76BD6F773737373737B73BDE773CEEF7B8C947BC6E784848484BDD684C6DE8C8C8C8CCEE79494949C9C9C9CBDCEA5A5A5ADADADADB5B5ADDEF7B5B5B5B5E7FFBDBDBDBDE7F7BDEFFFC6C6C6CECECED6D6D6D6DEDED6E7EFD6EFF7DEDEDEDEE7E7E7E7E7E7EFEFEFEFEFEFF7FFEFFFFFF7F7F7F7F7FFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100004A002C00000000100010000008C80095081402A460411E3B74E80022B061C21D3C800821C8E3468D8602811C392284878E1A2F52B8B888110811254780E89011C3C50B920D4DA204B2A346CB971895409C59B3A58B1B398300A9C18244870D183A8C88915366CA9E2E61663CF95405070D1A5614C1C8832A9016150E8870A020C456813565BC3851C18081221C0C24B0119348121E30121480800302010226EAA25CF177000D0903065C107884C711943E06086080E38300012518BF704CD04200061F3204407016658C11A83D3C00C09A820F250101003B00
END

/* Updating System Permissions Icon for Export */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 6

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 6

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000009C3100FF316363636B6B6B7373738484849494949C9C9CADADADB5C6E7BDBDBDC6C6C6CECECED6DEEFE7EFF7EFEFEFF7F7F7F7FFFFFF4242FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000013002C000000001000100000088B00271C1848B0A0418109122A5CA8F080C00610234A6C20A181C30310326ADCA8F1A2C18F04053A1839A1A4C99307060858C9B2254B8102264078B00081819800724E887940C0839F407106C8C9530083050A0E1420C012C050003065D2CC49D5E9D39E401F580DC0B52BD1A349ABE6E49AB3A7D49A374B3AD56976265A0107D60220A8D2E58009544F0604003B00
END

/* Updating System Permissions Icon for Filters */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 14

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 14

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700007B529C8452A58C63AD8C63B5946BB59473B59C73B59C73BD9C7BBDA57BBDA57BC6A584C6AD84C6AD8CC6AD8CCEB58CCEB594CEB59CCEBD9CD6BDA5D6C6A5D6C6ADDEC6B5DECEADDECEB5DECEBDDECEBDE7D6BDE7D6C6E7DEC6E7DECEE7DECEEFDED6EFE7D6EFE7DEEFEFDEF7EFE7F7F7EFF7F7EFFFF7F7FFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000029002C000000001000100000089C005308AC500183871202132A4C813081000219449C58A890848803020A34F02002054581244A4040D060C28610223E4A24C14142050E213C78A0E80185880E1E0C861451612187121826442038210409970A234E80B0A0008103080C1028A010C386080D3006008020E5C20A131A2420202002C28F141A2C588061E2C7141E3834E0F03621D0B0751362C860E06C5D111B4AE615B8A14283C1290202003B00
END

/* Updating System Permissions Icon for Global Add */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 7

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 7

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000009C186363636B6B6B7373737B7B7B8484848C8C8C9494949C9C9CA5A5A5ADADADB5C6E7BDBDBDC6C6C6CECECED6D6D6D6DEEFDEDEDEE7EFF7EFEFEFF7F7F7F7FFFFFF4242FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000017002C000000001000100000089C002F201848B0A041810B122A5CA81081400810234A84600182430414326ADCA8F1A2C18F04054A1879A1A4C99308040458C9B2254B81012E5098D020C181000D080430B0F201820013820A0D208081C90303783A68C000418101010E94045092814A993413148839D5E44AA141190CB800A06CD90B3A97367DCA95ACD798336BF234EAF68258B8596F0A4860328100032A5D0A78B015708108170202003B00
END

/* Updating System Permissions Icon for Global Delete */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 9

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 9

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700006363636B63636B6B6B7363637373737B5A5A848484944A4A9494949C9C9CA53939A58C8CAD3131AD3939AD8484ADADADB56363B56B6BB5C6E7BD5252BD6B6BBDBDBDC62929C65A5AC6C6C6CECECED61818D62121D6DEEFDE1818DE2121E71818E7EFF7EFEFEFF70000F71010F79494F7F7F7F7FFFFFF0000FF0808FF1010FF1818FF2121FF3131FF3939FF4242FF5252FF6363FF6B6BFF7373FF7B7BFF9C9CFFDEDEFFE7E7FFEFEFFFF7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000039002C00000000100010000008B60073241848B0A0418112122A5CA83081400E10234AE4E08283C30425326ADCA8F1A2C18F0405821899A3A4C99309040058B9324003050158021038B344080C0F20AC402163668E02311C000841942889192C5234C85103458A08003260A890C000810D30568868F061848C04356F3E401040860B1529566CB801B668D1992F5AA0403113AC54AA56075838BBA28506B6617122F01063278D162B6280CD6113E704A4282E24583062440495320FEC9C61D2C68C1A0101003B00
END

/* Updating System Permissions Icon for Global Update */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 8

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 8

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F7000029947B31313131947B42BDB552BDA55ABDA563636363948C6B6B636B6B6B7373737B7B6B7B7B738484849494949C9C9CA59C8CAD9442AD945AADADADB58429B58C31B59452B5A542B5B5ADB5C6E7BD8C29BD8C31BDBDBDC6A56BC6C6C6CEA539CECECED69418D69C18D69C21D69C29D6DEEFDE9C31DEA529DEA531DEAD4ADEB55AE7AD31E7B531E7B552E7BD63E7EFF7EFB539EFBD39EFE7E7EFEFEFF7F7EFF7F7F7F7FFFFFF4242FFEFADFFEFB5FFF7B5FFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100003C002C00000000100010000008AC00793C1848B0A0418119122A5CA8F081C01210234A2C71A384C30735326ADCA8F1A2C18F0405BE18C9A3A4491E350AB8B090C080CB97300D0C187021820194333C4C7070930782141F080010606086D1A337696C20B122848003064078E0F0A0810203083698801182028D07376BE4DC89400509182354201058F4E80C191598865041E32658A9541B482031A22B02976C717AC090038709BA27C10A8680234707902D5F0658C0E0E4C98000003B00
END

/* Updating System Permissions Icon for Import */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 10

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 10

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000009C3100FF31089C31636363636B636B6B6B7373738484849494949C9C9CA5ADA5ADADADB5C6E7BDBDBDC6C6C6CECECED6DEEFE7EFF7EFEFEFF7F7F7F7FFFFFF4242FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000016002C0000000010001000000889002D241848B0A041810C122A5CA83081400810234A84500182C30413326ADCA8F1A2C18F0405461869A1A4C993090A0C58C9B2254B81032C4C90E04001829802725A88996080849F4007E40C206027CC070E1A24386060A880950460CAA4A96028D19C587D02B51AA0ABD7A748951EB08A35EBD49A089CA28C3913AD5001440BAA7459C002D6930101003B00
END

/* Updating System Permissions Icon for Mail Merge */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 11

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 11

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700006363636B6B6B7373737B7B7B8484848C8C8C9494949C9C9CA5A5A5ADADADB5C6E7BDBDBDC6C6C6CECECED6D6D6D6DEEFDEDEDEE7EFF7EFEFEFF7F7F7F7FFFFFF0000FF4242FF4A4AFF7373FF8484FFADADFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100001B002C00000000100010000008A200371C1848B0A041810A122A5CA8F080C00710234A7C60E181C30313326ADCA8F1A2C18F04054618B9A1A4C99307020058C9B2254B8100364C90C0008101000C060028B0D2C1010012820A05106081490302783660B0E0000101000C9CDCB040A54C9A0808C4AC80614385AF2B85065D20A082860A1B32A0D5B9B4E9D3AD172AA8DDB0F26A4D9E464F928D39B3E6CD00084C22085040A5CB000EB416260061434000003B00
END

/* Updating System Permissions Icon for Match Reports */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 23

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 23

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700000852E72929AD3131AD3939AD3939B53994004242B5424AAD4A4AB54A4ABD5252BD5A5ABD6363BD6363C66B6BC67373CE7B7BCE8484CE84A5FF8C8CD69418319494D69C9CD6AD0808ADADDEB5B5E7BDBDE7BDE79CCECEEFCED6D6CEDEFFD6D6EFE75252E7F7D6EFEFF7EFEFFFF7B5B5FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000025002C00000000100010000008A50001081C48906089120024285C28C1834387000E169C38F0E0C10218336AD4787183C710203B3010C0A043880217371638B0A0828203182FC89C3933C0840A10041CBC00A227081240491830C0C000829D3469521810C0C0078B502D72406074000211512D8A40D0A04285080614643DF8408185AF111C08C09055C4809B132240706034AB860115E2CE6DA04067540E02B83A70B0008100BF513330187058808105220202003B00
END

/* Updating System Permissions Icon for Menu */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 37

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 37

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700000000638C8C8C94949494949C949C9C9C9C9C9C9CA5A5A5A5A5BDD6ADADADADC6DEADC6E7C6C6C6CE0000CECECEE7EFEFE7EFF7EFEFEFF7F7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000013002C00000000100010000008640027081C48B0A0C18304191458C8B02143030E02205040B1A245050D10081800A0A347001715682C00D222000828533E1840F2220083121696ACF8B260CC962661CA74A91327CD9E3329D62478F3A34783110A304870A0A9D3A74D0B384048B56AC18000003B00
END

/* Updating System Permissions Icon for Orders */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 16

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 16

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000000000005AE73173EF4242425252526363636394EF7373737BADF784848494B5F7B50000BD0000C64242CE0000CE2121D60000DE0000DE2121DE5252E70000E72121E72929E77373E79C9CEF0000EF7373EF9C9CEFB5B5EFEFEFEFF7FFF77373F79C9CF7C6C6F7CECEF7F7F7F7F7FFFFCECEFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000026002C0000000010001000000893004D082C9141844081004C7428D0E1A0890F1434384C4800000187152C4898980080C704024544D800C120C29307354030E1E0C2C18426600A9400C18183072F519A10E12084090E0C4CC29439C10106131816344029D3A1D3A10703489D1A806954011E4C180860C0AA400502B776F57A50EC449D5AB99AF080E08047006ECBAAF520A0EA0000036246A52A55E18086300302003B00
END

/* Updating System Permissions Icon for Outlook Calendar Queue */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 40

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 40

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700000084CE008CD60094DE009CDE00A5E700ADE700B5EF108CCE108CD610ADE710B5EF10BDEF1894D6189CDE18BDEF299CCE29B5E729B5EF31ADE731C6F739BDEF4294C642A5D64AA5CE52ADDE52CEF75AB5DE5AC6EF5ACEF763737B638C9C6394AD63CEF76B7B846B8C9C6B8CA56B94A56B94AD6B94B56B9CC66BBDE76BBDEF6BD6F773848473848C739CB573A5BD73A5C673ADC67BADC67BADCE7BB5CE7BBDDE848C9484949C849CA584ADCE84B5D684C6E78C949C8C9C9C8C9CA58CA5B58CBDD68CBDDE8CC6DE949CA594A5A594C6DE94C6E794CEE794DEF79CA5AD9CADAD9CADB59CB5C69CCEE79CCEEF9CD6EF9CDEF7A5ADB5A5BDC6A5C6DEA5D6EFA5DEF7ADB5B5ADB5BDADDEF7ADE7F7B5DEF7B5E7F7BDC6CEBDE7F7BDEFFFC6CECEC6CED6CECED6CED6D6CED6DED6DEE7D6E7F7D6EFF7DEE7E7DEE7EFDEE7F7E7E7EFE7EFEFE7F7FFEFEFEFEFEFF7EFF7F7EFFFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000071002C00000000100010000008F500DB9C194870A0998366C68C61234586C3181063BC7041D1458B164B4E64D1C231CB952B54A64C71D28449110F2798A86462C4089120407EFCC8810387870F2572E624C173844F11404574D8C1A3470F2142860C41C234491228363AC471E3068ED5AB58E1444D33C64C9AAF6AD8B43958150E8F1050B67C591B860D1B2A1B326438B2A6470830426CD8E851838D130A0654385000E246883863BC6C19D3830C850205DC6C289060C6E13466BE84A991250101085C200C18A061459CB06AC20CA1315AC0150902043C5881F0A0151F020230E0822240800A4B9922790AE50200062830003810258EF3E7CF751C0000C00299380101003B00
END

/* Updating System Permissions Icon for Payroll Transfer */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 41

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 41

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700004A524A5252525284525A5A5A635A5263635A6B6B636B6B6B736B6373736B7B7B6B847B7384847B84947B8C844A8C84528C845A8C84638C847B948C52948C5A948C6394948494A5849C8C529C8C5A9C8C639C94639C946B9C94849C9C6B9C9C739C9C849C9C8CA58C42A5945AA5946BA59C6BA59C73A59C7BA59C84A59C94A5A573AD9439AD9452AD945AAD9C52AD9C63ADA573ADA57BADA584B59439B5AD84B5AD8CB5AD9CB5B584B5B58CB5B59CBDB584BDB58CBDB594BDB5A5BDBD94BDBD9CC6AD63C6B573C6B584C6BD94C6BD9CC6BDA5C6C6A5C6C6ADCEC6A5CEC6ADD6C6ADD6CEADD6CEB5D6CEBDD6D6B5DECEA5DED6BDDEDECEE7E7DEEFDEBDEFE7C6EFE7D6EFE7DEEFEFE7F7EFDEF7F7EFF7F7F7FFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100005C002C00000000100010000008F100B9089462A448942C02132ACCF283020D191B1E5038C1238A96844C2A782812E28082114184BCE020D104130A392C0C308002C8142A23080048B0A0C303073D423070F0044B90080D045C286060C7060726762499C16506080606020440B0630885071BAE6618C282895726368658BD59A40912113526608D5183080C0A2AAE3EA8A0624495154978A8C85A62C7839B4C88D0A0A035060F23498CD2F84BE1C389155C56F83051E1EF901D242838E0F1018283091860D048E2D56A899B44BCB2B022A20284BF10E8D238518424851248BA2E095C83068F0D52046E9152E4C455138799F0307145A1C22B4C6844642E3020003B00
END

/* Updating System Permissions Icon for Picklists */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 15

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 15

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000317BDE3984DE3984E739BD3142AD6342BD3142BD3942C6314A8CE74A9CF74ABD424ACE3952BD4A52C64A52CE425A6B845A94E75A9CE75AA5F75AA5FF5AB58C5AC6525ACE4A639CEF63ADF763ADFF63B5BD63CE5A63D65A6BA5EF6BADF76BADFF6BB5FF6BBD846BC67B6BCE636BD65A6BD66373B5FF73CE6B73D66373D66B7BADEF7BB5DE7BB5F77BBDCE7BCE737BD67384ADEF84B5F784BDFF84D67B84DE7384DE7B8CB5EF8CC6FF8CD6848CDE7B8CDE8494BDF794C6F794C6FF94DE8C9CCEFF9CDE949CE794A5D6FFAD1818ADD6FFADDEFFADE7A5B51818B5394AB54A29B5B5CEB5CEEFB5D6FFBD2121BD4242BDDEFFBDE7B5BDE7FFBDEFB5C68C39C6DEFFC6E7FFC6EFBDCE7329CE9C39CEE7FFCEEFC6D66B29D69431D6AD39D6E7FFD6EFCED6EFFFD6F7CEDE7B29DEE7FFDEEFFFE7AD29E7ADADE7EFEFE7F7FFEFB529EFB531EFBD42EFC642EFC64AEFD684EFF7E7F7D65AF7D663F7E79CF7F7F7F7F7FFF7FFF7F7FFFFFFE7A5FFEFB5FFEFBDFFF7DEFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100007B002C00000000100010000008F800F7081CB8870E0C040010D8A033700E0E060C70CC8100218B1D3214079E98A145CB8C1001B2FCF0F0234B002F020D7C3970000A81005430C83499456005232C718810C00483070F540430DCF345018D0D2EBEECB8D0E3C712244DCC08DC0064C6062D2754C420E2654B93AF66A054B042826A8B0E4FBC942923E6EB10173372A02871224216B571E494B932C4898219284854A0B0834C99BC6BDA6049A2A4800F070E0C68C08B47CD1A385388EC503062C60C172B0EE751C326CE942C4F58BC99D1A04081B57AD4B431EDE5C7930404F794B153C68DE92C426E08C14DB0CB993358A654A9F24386040FB91F60492E410206133C9EEC0908003B00
END

/* Updating System Permissions Icon for Record Profile */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 34

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 34

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700004A00A55200A55200AD5A00AD6300AD6B08B57310B57318B57331B57B21BD7B39BD8421BD8429BD8439BD8C31BD8C39C68C42C69442C6944AC69C52C69C52CEA563CEA56BCEA56BD6AD6BD6AD73D6B57BD6B584D6B584DEBD84D6BD8CD6BD8CDEC694DEC69CDEC69CE7CE9CDECEA5E7CEADE7D6ADE7D6B5E7D6B5EFDEBDEFDEC6EFE7CEEFE7D6F7E7DEF7EFE7F7F7EFF7F7EFFFF7F7FFFFF7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000033002C00000000100010000008AB0067081C18838487832012860831702008151D1A0A64D8B0C489101B5C6874F1620608811908101830204000000854A8F438C3C28A0225242A98F142C50C0F1B564C301043C3860D0C67D6BCC98041820333360854DA40A54D0F0D580C549A744683175893426858E1A7D2AD5325486CC880638CA413C60E2C10E2C48A191AD2CEC0E895A1810F6E676490ABF68085BF19F6CA585A7746028913549C1DCB00F184C790254478D040624000003B00
END

/* Updating System Permissions Icon for Standard Reports */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 13

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 13

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F700004242426B6B6B7373737B7B7B8484848C8C8C9494949C9C9CA5A5A5ADADADB5B5B5BDBDBDC6C6C6CECECED6D6D6DEDEDEE7E7E7EFEFEFF7F7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000013002C00000000100010000008CC0027088CD02041810105164410C8D0C18203050E2C7040104100030D182264104142848F1F213080C84082830204124000C912428303060A286880C08082952C233C4810F340039D08082888E0D2E0000223090C30B0F28181014B13381049C040470504103C201A0100000308961228C011024F045F0DA8350040A90288431F2CC0D8A06E83B50416BC14A0E0EB5D0518D91670A0B68103B4811330B80B20C081070D96120E0CF8AE0109737DCE2D40D7EE538112181838C0C041DAB50024309C20A14101CE05BC5E161810003B00
END

/* Updating System Permissions Icon for Succession Planning */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 38

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 38

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F70000009C3100FF312929AD3939AD3939B54242B54A4AB54A4ABD5252BD5A5ABD6363BD6363C66B6BC67373C67B7BCE8484CE8C8CD69494D69C9CD6A5A5DEADADDEB5B5E7BDBDE7C6C6E7C6C6EFD6D6EFE7E7F7EFEFF7EFEFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100001D002C0000000010001000000891003B081CD84143060D04130E84A0A061048504272490600142020A1005326850A1A30206193B2848F0000204040A42462880604182021242766030A06603991B1E203000A067460D0C183C780020804F8541214CA030A1E8D181127B4A2D6A140041050DA806D8CAF56806040EA6F6DCFA940282A00F1C38C85A15A3400B061AA60D5B55C085882F0BE8956A6082C08000003B00
END

/* Updating System Permissions Icon for Workflow */
SELECT @iRecCount = count(*)
FROM ASRSysPermissionCategories
WHERE categoryID = 42

IF @iRecCount > 0 
BEGIN
	/* The record exists, so update it with new icon. */

	SELECT @ptrval = TEXTPTR(picture) 
	FROM ASRSysPermissionCategories
	WHERE categoryID = 42

	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x47494638396110001000F7000000BD0008AD0010BD1018B50821C62131CE3142CE424ACE4A52D6525AAD2963D6636B840873DE737B7B0084DE848CE78C9CBD009CE79CA5E7A5ADEFADBD9C42BDA542BDB56BBDBD8CBDE7ADC69C21C6A531C6A563C6B510C6B552C6B563C6BD52C6CE8CC6EFC6CE8C10CEAD7BCEC618CEC67BCECE10CECE7BCECE8CCEF7CED68408D6AD08D6CE08D6F7D6DE8C00DEE700DEF7DEE79C00E7BD00E7F7E7EFAD00EFDE00EFEF00EFFFEFF7EFE7F7F7E7F7FFF7FF9C00FFA500FFAD00FFB500FFBD00FFC600FFCE00FFD600FFDE00FFE700FFEF00FFF700FFFF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F90401000048002C00000000100010000008C00091081C3810C511130407DE8870A00080012F8E80488864820000180140380241400A821330228810A2C4111B181C781C4800400481386C1CF920D001028230069230C242210085046E202951844807810632DE540040018E1A44860891A1C1024601306064E4205548102040323800E00049528C093C5458F1E3070D1C376E0EFC3990828F1E1B282229F052600C1E220446004060A0CA073736F0D8712142D295031D0270B1A34146031F09DE70B060878B00061864D6AB62C708BD480202003B00
END

/* ------------------------------------------------------------- */
PRINT 'Step 15 - System Locking'

	----------------------------------------------------------------------
	-- sp_ASRLockCheck
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRLockCheck]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRLockCheck]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRLockCheck]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[sp_ASRLockCheck] AS
		BEGIN
		
			SET NOCOUNT ON
		
			DECLARE @sSQLVersion int
			SELECT @sSQLVersion = dbo.udfASRSQLVersion()
		
			IF @sSQLVersion >= 9 AND APP_NAME() <> ''HR Pro Workflow Service'' AND APP_NAME() <> ''HR Pro Outlook Calendar Service''
			BEGIN
				CREATE TABLE #tmpProcesses 
					(HostName varchar(100)
					,LoginName varchar(100)
					,Program_Name varchar(100)
					,HostProcess int
					,Sid binary(86)
					,Login_Time datetime
					,spid int
					,uid smallint)
				INSERT #tmpProcesses EXEC dbo.[spASRGetCurrentUsers]
		
				SELECT ASRSysLock.* FROM ASRSysLock
				LEFT OUTER JOIN #tmpProcesses syspro 
					ON ASRSysLock.spid = syspro.spid AND ASRSysLock.login_time = syspro.login_time
				WHERE priority = 2 OR syspro.spid IS NOT NULL
				ORDER BY priority
		
				DROP TABLE #tmpProcesses
			END
			ELSE
			BEGIN
				SELECT ASRSysLock.* FROM ASRSysLock
				LEFT OUTER JOIN master..sysprocesses syspro 
					ON asrsyslock.spid = syspro.spid AND asrsyslock.login_time = syspro.login_time
				WHERE Priority = 2 OR syspro.spid IS NOT NULL
				ORDER BY Priority
			END
		
		END'

	EXECUTE (@sSPCode_0)




/* ------------------------------------------------------------- */
PRINT 'Step 16 - Updating support contact details'

delete from asrsyssystemsettings
where [Section] = 'support' and [SettingKey] = 'email'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('support', 'email', 'service.delivery@coasolutions.com')




/* ------------------------------------------------------------- */
/* ------------------------------------------------------------- */

/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
     INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
    OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
    OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
    AND (sysusers.name = 'dbo')
--IF (@@ERROR <> 0) goto QuitWithRollback

OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
    IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'FN'
    BEGIN
        SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
        --IF (@@ERROR <> 0) goto QuitWithRollback
    END
    ELSE
    BEGIN
        SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
        --IF (@@ERROR <> 0) goto QuitWithRollback
    END

    FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects
/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '4.0')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '4.0.0')

delete from asrsyssystemsettings
where [Section] = 'ssintranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('ssintranet', 'minimum version', '4.0.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.4.0')

delete from asrsyssystemsettings
where [Section] = '.NET Assembly' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('.NET Assembly', 'minimum version', '4.0.0')

delete from asrsyssystemsettings
where [Section] = 'outlook service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('outlook service', 'minimum version', '4.0.0')

delete from asrsyssystemsettings
where [Section] = 'workflow service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('workflow service', 'minimum version', '4.0.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v4.0')


SELECT @NVarCommand = 
	'IF EXISTS (SELECT * FROM dbo.sysobjects
			WHERE id = object_id(N''[dbo].[sp_ASRLockCheck]'')
			AND OBJECTPROPERTY(id, N''IsProcedure'') = 1)
		GRANT EXECUTE ON sp_ASRLockCheck TO public'
EXEC sp_executesql @NVarCommand


SELECT @NVarCommand = 'USE master
GRANT EXECUTE ON sp_OACreate TO public
GRANT EXECUTE ON sp_OADestroy TO public
GRANT EXECUTE ON sp_OAGetErrorInfo TO public
GRANT EXECUTE ON sp_OAGetProperty TO public
GRANT EXECUTE ON sp_OAMethod TO public
GRANT EXECUTE ON sp_OASetProperty TO public
GRANT EXECUTE ON sp_OAStop TO public
GRANT EXECUTE ON xp_StartMail TO public
GRANT EXECUTE ON xp_SendMail TO public
GRANT EXECUTE ON xp_LoginConfig TO public
GRANT EXECUTE ON xp_EnumGroups TO public'
--EXEC sp_executesql @NVarCommand

SELECT @NVarCommand = 'USE ['+@DBName + ']
GRANT VIEW DEFINITION TO public'
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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v4.0 Of HR Pro'
