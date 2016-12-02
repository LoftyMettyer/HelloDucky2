
/* --------------------------------------------------- */
/* Update the database from version 3.5 to version 3.6 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(8000),
	@iSQLVersion numeric(3,1),
	@NVarCommand nvarchar(4000),
	@sObject sysname,
    @sObjectType char(2),
	@ptrval binary(16)

DECLARE @sSQL varchar(8000)
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
IF (@sDBVersion <> '3.5') and (@sDBVersion <> '3.6')
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
PRINT 'Step 1 of X - Modifying Workflow tables'

	/* SRSysWorkflowElements - Add new Web Form TimeoutExcludeWeekend column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'TimeoutExcludeWeekend'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							TimeoutExcludeWeekend [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.TimeoutExcludeWeekend = 0
							WHERE ASRSysWorkflowElements.TimeoutExcludeWeekend IS NULL'
		EXEC sp_executesql @NVarCommand
	END


	/* ASRSysWorkflowElementsItems - Add new WorkflowItems VerticalOffset column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'VerticalOffset'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							VerticalOffset [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElementItems
							SET ASRSysWorkflowElementItems.VerticalOffset = 0
							WHERE ASRSysWorkflowElementItems.VerticalOffset IS NULL'
		EXEC sp_executesql @NVarCommand
	END


	/* ASRSysWorkflowElementsItems - Add new WorkflowItems VerticalOffsetBehaviour column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'VerticalOffsetBehaviour'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							VerticalOffsetBehaviour [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElementItems
							SET ASRSysWorkflowElementItems.VerticalOffsetBehaviour = 0
							WHERE ASRSysWorkflowElementItems.VerticalOffsetBehaviour IS NULL'
		EXEC sp_executesql @NVarCommand
	END


	/* ASRSysWorkflowElementsItems - Add new WorkflowItems HorizontalOffset column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'HorizontalOffset'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							HorizontalOffset [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElementItems
							SET ASRSysWorkflowElementItems.HorizontalOffset = 0
							WHERE ASRSysWorkflowElementItems.HorizontalOffset IS NULL'
		EXEC sp_executesql @NVarCommand
	END


	/* ASRSysWorkflowElementsItems - Add new WorkflowItems HorizontalOffsetBehaviour column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'HorizontalOffsetBehaviour'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							HorizontalOffsetBehaviour [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElementItems
							SET ASRSysWorkflowElementItems.HorizontalOffsetBehaviour = 0
							WHERE ASRSysWorkflowElementItems.HorizontalOffsetBehaviour IS NULL'
		EXEC sp_executesql @NVarCommand
	END


	/* ASRSysWorkflowElementsItems - Add new WorkflowItems HeightBehaviour column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'HeightBehaviour'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							HeightBehaviour [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElementItems
							SET ASRSysWorkflowElementItems.HeightBehaviour = 0
							WHERE ASRSysWorkflowElementItems.HeightBehaviour IS NULL'
		EXEC sp_executesql @NVarCommand
	END


	/* ASRSysWorkflowElementsItems - Add new WorkflowItems WidthBehaviour column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'WidthBehaviour'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							WidthBehaviour [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElementItems
							SET ASRSysWorkflowElementItems.WidthBehaviour = 0
							WHERE ASRSysWorkflowElementItems.WidthBehaviour IS NULL'
		EXEC sp_executesql @NVarCommand
	END
	
	/* ASRSysWorkflowElementsColumns - Add new Workflow Columns CalcID column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementColumns', 'U')
	AND name = 'CalcID'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementColumns ADD 
							CalcID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElementColumns
							SET ASRSysWorkflowElementColumns.CalcID = 0
							WHERE ASRSysWorkflowElementColumns.CalcID IS NULL'
		EXEC sp_executesql @NVarCommand
	END


	/* ASRSysSystemSettings - Update Webform controls with new property defaults */
	IF NOT EXISTS(SELECT ASRSysSystemSettings.SettingValue FROM ASRSysSystemSettings 
					WHERE ASRSysSystemSettings.Section = 'workflow' 
					AND ASRSysSystemSettings.SettingKey = 'updatewfitemprops')
	BEGIN
		/* Buttons */
		UPDATE ASRSysWorkflowElementItems
		SET BackColor = -2147483633
			, ForeColor = -2147483630
		WHERE ItemType = 0
		
		/* Lines */
		UPDATE ASRSysWorkflowElementItems
		SET BackColor = -2147483632
			, Orientation = 1
		WHERE ItemType = 9			

		UPDATE ASRSysWorkflowElementItems
		SET Orientation = 0
		WHERE ItemType = 15	

		DELETE FROM ASRSysSystemSettings 
		WHERE [Section] = 'workflow' AND [SettingKey] = 'updatewfitemprops'
		INSERT INTO ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
		VALUES ('workflow', 'updatewfitemprops', 0)
	
		UPDATE ASRSysWorkflowElements
		SET WebFormDefaultFontSize = 8
		WHERE WebFormDefaultFontSize = 7
	END
/* ------------------------------------------------------------- */


/* ------------------------------------------------------------- */
PRINT 'Step 2 of X - Updating Export Definitions'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExportDetails')
	and name = 'ConvertCase'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExportDetails ADD [ConvertCase] [smallint] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysExportDetails SET [ConvertCase] = 0'
		EXEC sp_executesql @NVarCommand
	END
/* ------------------------------------------------------------- */


/* ------------------------------------------------------------- */
PRINT 'Step 3 of X - Updating E-Mail Link Definitions'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysSSIntranetLinks')
	and (name = 'EMailAddress' or name = 'EMailSubject')

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD EMailAddress varchar(500) NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysSSIntranetLinks SET EMailAddress = '''''
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD EMailSubject varchar(500) NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysSSIntranetLinks SET EMailSubject = '''''
		EXEC sp_executesql @NVarCommand

	END

/* ------------------------------------------------------------- */
PRINT 'Step 4 of X - Overnight Job Stored Procedures'

	----------------------------------------------------------------------
	-- spASRSysOvernightTableUpdate
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSysOvernightTableUpdate]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSysOvernightTableUpdate]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRSysOvernightTableUpdate]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRSysOvernightTableUpdate]
		(
			@piTableName varchar(255),
			@piFieldName varchar(255),
			@piBatches int
		) 
		AS
		BEGIN
			SET NOCOUNT ON
		
			-- Create the progress table if it doesn''t already exist
			IF OBJECT_ID(''ASRSysOvernightProgress'', N''U'') IS NULL
				CREATE TABLE ASRSysOvernightProgress
					(TableName varchar(255)
					, RecCount int
					, IDRange varchar(255)
					, StartDate datetime
					, EndDate datetime
					, DurationMins int)
		
			DECLARE @lowid int,@highid int,@maxid int
			DECLARE @rowcount int, @start datetime
		
			DECLARE @sSQL nvarchar(4000)
			DECLARE @sParamDefinition nvarchar(4000)
		
			-- Determine the number of ID''s we''ll update in each batch
			IF ISNULL(@piBatches, 0) = 0
				SET @piBatches = 2000
			SET @lowid = 0 
			SET @highid = @lowid + @piBatches
			
			SET @sSQL = ''SELECT @maxid = ISNULL(MAX(ID),0) FROM '' + @piTableName
			SET @sParamDefinition = N''@maxid int OUTPUT''
			EXEC sp_executesql @sSQL, @sParamDefinition, @maxid OUTPUT
		
			WHILE 1=1
			BEGIN
				SET @start = GETDATE()
				
				-- Do the update
				SELECT @sSQL = ''UPDATE '' + @piTableName + '' SET '' + @piFieldName + '' = '' + @piFieldName
							+ '' WHERE ID BETWEEN @lowid AND @highid''
				SET @sParamDefinition = N''@lowid int, @highid int''
				EXEC sp_executesql @sSQL, @sParamDefinition, @lowid, @highid
		
				SET @rowcount = @@ROWCOUNT
		
				-- insert a record to this progress table to check the progress
				INSERT INTO ASRSysOvernightProgress 
					SELECT @piTableName
						, @rowcount
						, CAST(@lowid as varchar(255)) + ''-'' + CAST(@highid as varchar(255))
						, @start
						, GETDATE()
						, DATEDIFF(n, @start, GETDATE())
		
				SET @lowid = @lowid + @piBatches
				SET @highid = @lowid + @piBatches
		
				IF @lowid > @maxid
				BEGIN
					CHECKPOINT
					BREAK
				END
				ELSE
					CHECKPOINT
			END
		
			SET NOCOUNT OFF
		END'

	EXECUTE (@sSPCode_0)

/* ---------------------------------------------------------------------------------- */
PRINT 'Step 5 of X - Current Date on Server'

	-- [udfASRGetDate]
	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfASRGetDate]') AND sysstat & 0xf = 0)
		DROP FUNCTION [dbo].[udfASRGetDate]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRGetDate]
	(
	)
	RETURNS datetime
	AS
	BEGIN

		DECLARE @dtDate datetime

		SELECT TOP 1 @dtDate = convert(datetime, convert(varchar(20), last_batch, 101))
		FROM master..sysprocesses
		ORDER BY last_batch DESC

		RETURN @dtDate

	END'
	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */


/* ------------------------------------------------------------- */
PRINT 'Step 6 of X - Modifying Workflow stored procedures'

	----------------------------------------------------------------------
	-- spASRGetStoredDataActionDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetStoredDataActionDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetStoredDataActionDetails]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetStoredDataActionDetails]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetStoredDataActionDetails
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psSQL				varchar(8000)	OUTPUT, 
				@piDataTableID		integer			OUTPUT,
				@psTableName		varchar(8000)	OUTPUT,
				@piDataAction		integer			OUTPUT, 
				@piRecordID			integer			OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iPersonnelTableID			integer,
					@iInitiatorID				integer,
					@iDataRecord				integer,
					@sIDColumnName				varchar(8000),
					@iColumnID					integer, 
					@sColumnName				varchar(8000), 
					@iColumnDataType			integer, 
					@sColumnList				varchar(8000),
					@sValueList					varchar(8000),
					@sValue						varchar(8000),
					@sRecSelWebFormIdentifier	varchar(8000),
					@sRecSelIdentifier			varchar(8000),
					@iTempTableID				integer,
					@iSecondaryDataRecord		integer,
					@sSecondaryRecSelWebFormIdentifier	varchar(8000),
					@sSecondaryRecSelIdentifier	varchar(8000),
					@sSecondaryIDColumnName		varchar(8000),
					@iSecondaryRecordID			integer,
					@iElementType				integer,
					@iWorkflowID				integer,
					@iID int,
					@sWFFormIdentifier varchar(8000),
					@sWFValueIdentifier varchar(8000),
					@iDBColumnID int,
					@iDBRecord int,
					@sSQL nvarchar(4000),
					@sParam nvarchar(4000),
					@sDBColumnName nvarchar(4000),
					@sDBTableName nvarchar(4000),
					@iRecordID int,
					@sDBValue varchar(8000),
					@iDataType int, 
					@iValueType int, 
					@iSDColumnID int,
					@fValidRecordID	bit,
					@iBaseTableID	integer,
					@iBaseRecordID	integer,
					@iRequiredTableID	integer,
					@iRequiredRecordID	integer,
					@iDataRecordTableID	integer,
					@iSecondaryDataRecordTableID	integer,
					@iParent1TableID	int,
					@iParent1RecordID	int,
					@iParent2TableID	int,
					@iParent2RecordID	int,
					@iInitParent1TableID	int,
					@iInitParent1RecordID	int,
					@iInitParent2TableID	int,
					@iInitParent2RecordID	int,
					@iEmailID		int,
					@iType			int,
					@fDeletedValue	bit,
					@iTempElementID	integer,
					@iCount			integer,
					@iResultType	integer,
					@sResult		varchar(8000),
					@fResult		bit,
					@dtResult		datetime,
					@fltResult		float,
					@iCalcID		integer,
					@iSize		integer,
					@iDecimals	integer,
					@iTriggerTableID	int
						
				SET @psSQL = ''''
				SET @piRecordID = 0
			
				SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel''
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
					@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
					@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
					@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
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
				WHERE ASRSysWorkflowElements.ID = @piElementID
			
				SELECT @psTableName = tableName
				FROM ASRSysTables
				WHERE tableID = @piDataTableID
			
				IF @iDataRecord = 0 -- 0 = Initiator''s record
				BEGIN
					EXEC [dbo].['


	SET @sSPCode_1 = 'spASRWorkflowAscendantRecordID]
						@iPersonnelTableID,
						@iInitiatorID,
						@iInitParent1TableID,
						@iInitParent1RecordID,
						@iInitParent2TableID,
						@iInitParent2RecordID,
						@iDataRecordTableID,
						@piRecordID	OUTPUT
			
					IF @piDataTableID = @iDataRecordTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iDataRecordTableID)
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
						@piRecordID	OUTPUT
			
					IF @piDataTableID = @iDataRecordTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iDataRecordTableID)
					END
				END

				IF @iDataRecord = 1 -- 1 = Identified record
				BEGIN
					SELECT @iElementType = ASRSysWorkflowElements.type
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
						AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))
					
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
							AND IV.elementID = Es.ID
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
						WHERE IV.instanceID = @piInstanceID
					END

					SET @piRecordID = 
						CASE
							WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
							ELSE 0
						END
				
					SET @iBaseTableID = @iTempTableID
					SET @iBaseRecordID = @piRecordID
					EXEC [dbo].[spASRWorkflowAscendantRecordID]
						@iBaseTableID,
						@iBaseRecordID,
						@iParent1TableID,
						@iParent1RecordID,
						@iParent2TableID,
						@iParent2RecordID,
						@iDataRecordTableID,
						@piRecordID	OUTPUT

					IF @piDataTableID = @iDataRecordTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iDataRecordTableID)
					END
				END
			
				SET @fValidRecordID = 1
				IF (@iDataRecord = 0) OR (@iDataRecord = 1) OR (@iDataRecord = 4)
				BEGIN
					EXEC [dbo].[spASRWorkflowValidTableRecord]
						@iDataRecordTableID,
						@piRecordID,
						@fValidRecordID	OUTPUT

					IF @fValidRecordID = 0
					BEGIN
						-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
						EXEC [dbo].[spASRWorkflowActionFailed]
							@piInstanceID, 
							@piElementID, 
							''Stored Data primary record has been deleted or not selected.''

'


	SET @sSPCode_2 = '
						SET @psSQL = ''''
						RETURN
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
							@iSecondaryRecordID	OUTPUT
				
						IF @piDataTableID = @iSecondaryDataRecordTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iSecondaryDataRecordTableID)
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
							@iSecondaryRecordID	OUTPUT
				
						IF @piDataTableID = @iSecondaryDataRecordTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iSecondaryDataRecordTableID)
						END
					END

					IF @iSecondaryDataRecord = 1 -- 1 = Previous record selector''s record
					BEGIN
						SELECT @iElementType = ASRSysWorkflowElements.type
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sSecondaryRecSelWebFormIdentifier)))
				
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
								AND IV.elementID = Es.ID
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
							WHERE IV.instanceID = @piInstanceID
						END

						SET @iSecondaryRecordID = 
							CASE
								WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
								ELSE 0
							END
						
						SET @iBaseTableID = @iTempTableID
						SET @iBaseRecordID = @iSecondaryRecordID
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
							@iBaseTableID,
							@iBaseRecordID,
							@iParent1TableID,
							@iParent1RecordID,
							@iParent2TableID,
							@iParent2RecordID,
							@iSecondaryDataRecordTableID,
							@iSecondaryRecordID	OUTPUT

						IF @piDataTableID = @iSecondaryDataRecordTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iSecondaryDataRecordTableID)
						END'


	SET @sSPCode_3 = '
					END

					SET @fValidRecordID = 1
					IF (@iSecondaryDataRecord = 0) OR (@iSecondaryDataRecord = 1) OR (@iSecondaryDataRecord = 4)
					BEGIN
						EXEC [dbo].[spASRWorkflowValidTableRecord]
							@iSecondaryDataRecordTableID,
							@iSecondaryRecordID,
							@fValidRecordID	OUTPUT

						IF @fValidRecordID = 0
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC [dbo].[spASRWorkflowActionFailed] 
								@piInstanceID, 
								@piElementID, 
								''Stored Data secondary record has been deleted or not selected.''

							SET @psSQL = ''''
							RETURN
						END
					END
				END

				IF @piDataAction = 0 OR @piDataAction = 1
				BEGIN
					/* INSERT or UPDATE. */
					SET @sColumnList = ''''
					SET @sValueList = ''''

					DECLARE @dbValues TABLE (
						ID integer, 
						wfFormIdentifier varchar(1000),
						wfValueIdentifier varchar(1000),
						dbColumnID int,
						dbRecord int,
						value varchar(8000)
					)

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
						AND EC.valueType = 2
						
					DECLARE dbValuesCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT ID,
						wfFormIdentifier,
						wfValueIdentifier,
						dbColumnID,
						dbRecord
					FROM @dbValues
					OPEN dbValuesCursor
					FETCH NEXT FROM dbValuesCursor INTO @iID,
						@sWFFormIdentifier,
						@sWFValueIdentifier,
						@iDBColumnID,
						@iDBRecord
					WHILE (@@fetch_status = 0)
					BEGIN
						SET @fDeletedValue = 0

						SELECT @sDBTableName = tbl.tableName,
							@iRequiredTableID = tbl.tableID, 
							@sDBColumnName = col.columnName,
							@iDataType = col.dataType
						FROM ASRSysColumns col
						INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
						WHERE col.columnID = @iDBColumnID

						SET @sSQL = ''SELECT @sDBValue = ''
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN ''convert(varchar(8000),''
								ELSE ''convert(varchar(8000),''
							END
							+ @sDBTableName + ''.'' + @sDBColumnName
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN '', 101)''
								ELSE '')''
							END
							+ '' FROM '' + @sDBTableName 
							+ '' WHERE '' + @sDBTableName + ''.ID = ''

						SET @iRecordID = 0

						IF @iDBRecord = 0
						BEGIN
							-- Initiator''s record
							SET @iRecordID = @iInitiatorID
							SET @iParent1TableID = @iInitParent1TableID
							SET @iParent1RecordID = @iInitParent1RecordID
							SET @iParent2TableID = @iInitParent2TableID
							SET @iParent2RecordID = @iInitParent2RecordID

							SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
							FROM ASRSysModuleSetup
							WHERE moduleKey = ''MODULE_WORKFLOW''
							AND parameterKey = ''Param_TablePersonnel''
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
							-- Identified record
							SELECT @iElementType = ASRSysWorkflowElements.type, 
								@iTempElementID = ASRSysWorkflowElements.ID
							FROM ASRSysWorkflowElements
							WHERE ASRSysWorkflowElem'


	SET @sSPCode_4 = 'ents.workflowID = @iWorkflowID
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

						SET @fValidRecordID = 1
						
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
										SELECT @sDBValue = rtrim(ltrim(isnull(QC.columnValue , '''')))
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
											SELECT @sDBValue = rtrim(ltrim(isnull(IV.value , '''')))
											FROM ASRSysWorkflowInstanceValues IV
											WHERE IV.instanceID = @piInstanceID
												AND IV.columnID = @iDBColumnID
												AND IV.elementID = @iTempElementID

											SET @fValidRecordID = 1
											SET @fDeletedValue = 1
										END
									END
								END'


	SET @sSPCode_5 = '
							END

							IF @fValidRecordID = 0
							BEGIN
								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC [dbo].[spASRWorkflowActionFailed]
									@piInstanceID, 
									@piElementID, 
									''Stored Data column database value record has been deleted or not selected.''

								SET @psSQL = ''''
								RETURN
							END
						END

						IF @fDeletedValue = 0
						BEGIN
							SET @sSQL = @sSQL + convert(nvarchar(4000), @iRecordID)
							SET @sParam = N''@sDBValue varchar(8000) OUTPUT''
							EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT
						END

						UPDATE @dbValues
						SET value = @sDBValue
						WHERE ID = @iID
						
						FETCH NEXT FROM dbValuesCursor INTO @iID,
							@sWFFormIdentifier,
							@sWFValueIdentifier,
							@iDBColumnID,
							@iDBRecord
					END
					CLOSE dbValuesCursor
					DEALLOCATE dbValuesCursor
			
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
			
					OPEN columnCursor
					FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID, @iCalcID, @iSize, @iDecimals
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @iValueType = 2 -- DBValue - get here to avoid collation conflict
						BEGIN
							SELECT @sValue = dbV.value
							FROM @dbValues dbV
							WHERE dbV.ID = @iSDColumnID
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
								0

							IF @iColumnDataType = 12 SET @sResult = LEFT(@sResult, @iSize) -- Character
							IF @iColumnDataType = 2 -- Numeric
							BEGIN
								IF @fltResult >= power(10, @iSize - @iDecimals) SET @fltResult = 0
								IF @fltResult <= (-1 * power(10, @iSize - @iDecimals)) SET @fltResult = 0
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
									ELSE convert(varchar(8000), @sResult)
								END
						END

						IF @piDataAction = 0 
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
						'


	SET @sSPCode_6 = '			WHEN LEN(@sValueList) > 0 THEN '',''
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
									WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''''''' + replace(isnull(@sValue, ''''), '''''''', '''''''''''') + '''''''' -- 12 = varchar, -1 = working pattern
									WHEN @iColumnDataType = 11 THEN
										CASE 
											WHEN (upper(ltrim(rtrim(@sValue))) = ''NULL'') OR (@sValue IS null) THEN ''null''
											ELSE '''''''' + replace(@sValue, '''''''', '''''''''''') + '''''''' -- 11 = date
										END
									WHEN LEN(@sValue) = 0 THEN ''0''
									ELSE isnull(@sValue, 0) -- integer, logic, numeric
								END
						END

						DELETE FROM ASRSysWorkflowInstanceValues
						WHERE instanceID = @piInstanceID
							AND elementID = @piElementID
							AND columnID = @iDBColumnID

						INSERT INTO ASRSysWorkflowInstanceValues
							(instanceID, elementID, identifier, columnID, value, emailID)
							VALUES (@piInstanceID, @piElementID, '''', @iColumnID, @sValue, 0)
			
						FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID, @iCalcID, @iSize, @iDecimals
					END
			
					CLOSE columnCursor
					DEALLOCATE columnCursor
			
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
								+ @sIDColumnName
				
							SET @sValueList = @sValueList
								+ CASE
									WHEN LEN(@sValueList) > 0 THEN '',''
									ELSE ''''
								END
								+ convert(varchar(8000), @piRecordID)

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
									+ @sSecondaryIDColumnName
							
								SET @sValueList = @sValueList
									+ CASE
										WHEN LEN(@sValueList) > 0 THEN '',''
										ELSE ''''
									END
									+ convert(varchar(8000), @iSecondaryRecordID)
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
								+ '' VALUES('' + @sValueList + '')''
						END
						ELSE
						BEGIN
							/* UPDATE. */
							SET @psSQL = ''UPDATE '' + @psTableName
								+ '' SET '' + @sColumnList
								+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(8000), @piRecordID)
						END
					END
				END
			
				IF @piDataAction = 2
				BEGIN
					/* DELETE. */
					SET @psSQL = ''DELETE FROM '' + @psTableName
						+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(8000), @piRecordID)
				END	

				IF (@piDataAction = 0) -- Insert
				BEGIN
				'


	SET @sSPCode_7 = '	SET @iParent1TableID = isnull(@iDataRecordTableID, 0)
					SET @iParent1RecordID = isnull(@piRecordID, 0)
					SET @iParent2TableID = isnull(@iSecondaryDataRecordTableID, 0)
					SET @iParent2RecordID = isnull(@iSecondaryRecordID, 0)
				END
				ELSE
				BEGIN	-- Update or Delete
					exec [dbo].[spASRGetParentDetails]
						@piDataTableID,
						@piRecordID,
						@iParent1TableID	OUTPUT,
						@iParent1RecordID	OUTPUT,
						@iParent2TableID	OUTPUT,
						@iParent2RecordID	OUTPUT
				END

				UPDATE ASRSysWorkflowInstanceValues
				SET ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID, 
					ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
					ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID, 
					ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
				WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceValues.elementID = @piElementID
					AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
					AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0

				IF (@piDataAction = 2) -- Delete
				BEGIN
					DECLARE curColumns CURSOR LOCAL FAST_FORWARD FOR 
					SELECT columnID
					FROM [dbo].[udfASRWorkflowColumnsUsed] (@iWorkflowID, @piElementID, 0)

					OPEN curColumns

					FETCH NEXT FROM curColumns INTO @iDBColumnID
					WHILE (@@fetch_status = 0)
					BEGIN
						DELETE FROM ASRSysWorkflowInstanceValues
						WHERE instanceID = @piInstanceID
							AND elementID = @piElementID
							AND columnID = @iDBColumnID

						SELECT @sDBTableName = tbl.tableName,
							@iRequiredTableID = tbl.tableID, 
							@sDBColumnName = col.columnName,
							@iDataType = col.dataType
						FROM ASRSysColumns col
						INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
						WHERE col.columnID = @iDBColumnID

						SET @sSQL = ''SELECT @sDBValue = ''
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN ''convert(varchar(8000),''
								ELSE ''convert(varchar(8000),''
							END
							+ @sDBTableName + ''.'' + @sDBColumnName
							+ CASE
								WHEN @iDataType = 12 THEN ''''
								WHEN @iDataType = 11 THEN '', 101)''
								ELSE '')''
							END
							+ '' FROM '' + @sDBTableName 
							+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(8000), @piRecordID)

						SET @sParam = N''@sDBValue varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT

						INSERT INTO ASRSysWorkflowInstanceValues
							(instanceID, elementID, identifier, columnID, value, emailID)
							VALUES (@piInstanceID, @piElementID, '''', @iDBColumnID, @sDBValue, 0)
								
						FETCH NEXT FROM curColumns INTO @iDBColumnID
					END
					CLOSE curColumns
					DEALLOCATE curColumns

					DECLARE curEmails CURSOR LOCAL FAST_FORWARD FOR 
					SELECT emailID,
						type,
						colExprID
					FROM [dbo].[udfASRWorkflowEmailsUsed] (@iWorkflowID, @piElementID, 0)

					OPEN curEmails

					FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID
					WHILE (@@fetch_status = 0)
					BEGIN
						DELETE FROM ASRSysWorkflowInstanceValues
						WHERE instanceID = @piInstanceID
							AND elementID = @piElementID
							AND emailID = @iEmailID

						IF @iType = 1 -- Column
						BEGIN
							SELECT @sDBTableName = tbl.tableName,
								@iRequiredTableID = tbl.tableID, 
								@sDBColumnName = col.columnName,
								@iDataType = col.dataType
							FROM ASRSysColumns col
							INNER JOIN ASRSysTables tbl ON col.tableID = tbl.tableID
							WHERE col.columnID = @iDBColumnID

							SET @sSQL = ''SELECT @sDBValue = ''
								+ CASE
									WHEN @iDataType = 12 THEN ''''
									WHEN @iDataType = 11 THEN ''convert(varchar(8000),''
									ELSE ''convert(varchar(8000),''
								END
								+ @sDBTableName + ''.'' + @sDBColumnName
								+ CASE
									WHEN @iDataType = 12 TH'


	SET @sSPCode_8 = 'EN ''''
									WHEN @iDataType = 11 THEN '', 101)''
									ELSE '')''
								END
								+ '' FROM '' + @sDBTableName 
								+ '' WHERE '' + @sDBTableName + ''.ID = '' + convert(varchar(8000), @piRecordID)

							SET @sParam = N''@sDBValue varchar(8000) OUTPUT''
							EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT
						END
						ELSE
						BEGIN
							EXEC [dbo].[spASRSysEmailAddr]
								@sDBValue OUTPUT,
								@iEmailID,
								@piRecordID
						END

						INSERT INTO ASRSysWorkflowInstanceValues
							(instanceID, elementID, identifier, columnID, value, emailID)
							VALUES (@piInstanceID, @piElementID, '''', 0, @sDBValue, @iEmailID)
								
						FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID
					END
					CLOSE curEmails
					DEALLOCATE curEmails
				END
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4
		+ @sSPCode_5
		+ @sSPCode_6
		+ @sSPCode_7
		+ @sSPCode_8)

	----------------------------------------------------------------------
	-- spASRGetWorkflowEmailMessage
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowEmailMessage]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowEmailMessage]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowEmailMessage]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRGetWorkflowEmailMessage]
							(
								@piInstanceID	integer,
								@piElementID	integer,
								@psMessage		varchar(8000)	OUTPUT, 
								@psMessage_HypertextLinks		varchar(8000)	OUTPUT, 
								@psHypertextLinkedSteps			varchar(8000)	OUTPUT, 
								@pfOK			bit	OUTPUT,
								@psTo			varchar(8000)
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
									@sTemp		varchar(8000),
									@sTableName		sysname,
									@sColumnName		sysname,
									@iRecordID		integer,
									@sSQL			nvarchar(4000),
									@sSQLParam		nvarchar(4000),
									@iCount		integer,
									@iElementID		integer,
									@superCursor		cursor,
									@iTemp		integer,
									@hResult 		integer,
									@objectToken 		integer,
									@sQueryString		varchar(8000),
									@sURL			varchar(8000), 
									@sEmailFormat		varchar(8000),
									@iEmailFormat		integer,
									@iSourceItemType	integer,
									@dtTempDate		datetime, 
									@sParam1	varchar(8000),
									@sDBName	sysname,
									@sRecSelWebFormIdentifier	varchar(8000),
									@sRecSelIdentifier	varchar(8000),
									@iElementType		integer,
									@iWorkflowID		integer, 
									@fValidRecordID	bit,
									@iBaseTableID	integer,
									@iBaseRecordID	integer,
									@iRequiredTableID	integer,
									@iRequiredRecordID	integer,
									@iParent1TableID	int,
									@iParent1RecordID	int,
									@iParent2TableID	int,
									@iParent2RecordID	int,
									@iInitParent1TableID	int,
									@iInitParent1RecordID	int,
									@iInitParent2TableID	int,
									@iInitParent2RecordID	int,
									@fDeletedValue		bit,
									@iTempElementID		integer,
									@iColumnID			integer,
									@iResultType	integer,
									@sResult		varchar(8000),
									@fResult		bit,
									@dtResult		datetime,
									@fltResult		float,
									@iCalcID		integer,
									@iSQLVersion	integer
											
								SET @pfOK = 1
								SET @psMessage = ''''
								SET @psMessage_HypertextLinks = ''''
								SET @psHypertextLinkedSteps = ''''
								SELECT @iSQLVersion = dbo.udfASRSQLVersion()
							
								exec [dbo].[spASRGetSetting]
									''email'',
									''date format'',
									''103'',
									0,
									@sEmailFormat		OUTPUT
					
								SET @iEmailFormat = convert(integer, @sEmailFormat)
								
								SELECT @sURL = parameterValue
								FROM ASRSysModuleSetup
								WHERE moduleKey = ''MODULE_WORKFLOW''
									AND parameterKey = ''Param_URL''
				
								IF upper(right(@sURL, 5)) <> ''.ASPX''
									AND right(@sURL, 1) <> ''/''
									AND len(@sURL) > 0
								BEGIN
									SET @sURL = @sURL + ''/''
								END
					
								SELECT @sParam1 = parameterValue
								FROM ASRSysModuleSetup
								WHERE moduleKey = ''MODULE_WORKFLOW''		
									AND parameterKey = ''Param_Web1''
								
								SET @sDBName = db_name()
					
								SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
									@iWorkflowID = ASRSysWorkflowInstances.workflowID,
									@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
									@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
									@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
									@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
								FROM ASRSysWorkflowInstances
								WHERE ASRSysWorkflowInstances.ID = @piInstanceID
							
								DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
								SELECT EI.caption,
									EI.itemType,
									EI.dbColumnID,
									EI.dbRecord,
									EI.wfFo'


	SET @sSPCode_1 = 'rmIdentifier,
									EI.wfValueIdentifier, 
									EI.recSelWebFormIdentifier,
									EI.recSelIdentifier, 
									EI.calcID
								FROM ASRSysWorkflowElementItems EI
								WHERE EI.elementID = @piElementID
								ORDER BY EI.ID
							
								OPEN itemCursor
								FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier, @iCalcID
								WHILE (@@fetch_status = 0)
								BEGIN
									SET @sValue = ''''
		
									IF @iItemType = 1
									BEGIN
										SET @fDeletedValue = 0
		
										/* Database value. */
										SELECT @sTableName = ASRSysTables.tableName, 
											@iRequiredTableID = ASRSysTables.tableID, 
											@sColumnName = ASRSysColumns.columnName, 
											@iSourceItemType = ASRSysColumns.dataType
										FROM ASRSysColumns
										INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
										WHERE ASRSysColumns.columnID = @iDBColumnID
							
										IF @iDBRecord = 0
										BEGIN
											-- Initiator''s record
											SET @iRecordID = @iInitiatorID
											SET @iParent1TableID = @iInitParent1TableID
											SET @iParent1RecordID = @iInitParent1RecordID
											SET @iParent2TableID = @iInitParent2TableID
											SET @iParent2RecordID = @iInitParent2RecordID
		
											SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
											FROM ASRSysModuleSetup
											WHERE moduleKey = ''MODULE_WORKFLOW''
											AND parameterKey = ''Param_TablePersonnel''
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
											-- Previously identified record.
											SELECT @iElementType = ASRSysWorkflowElements.type, 
												@iTempElementID = ASRSysWorkflowElements.ID
											FROM ASRSysWorkflowElements
											WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
												AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))
					
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
													AND IV.elementID = Es.ID
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
											'


	SET @sSPCode_2 = '	INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
													AND IV.identifier = Es.identifier
													AND Es.workflowID = @iWorkflowID
													AND Es.identifier = @sRecSelWebFormIdentifier
												WHERE IV.instanceID = @piInstanceID
											END
		
											SET @iRecordID = 
												CASE
													WHEN isnumeric(@sTemp) = 1 THEN convert(integer, @sTemp)
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
		
											IF @fValidRecordID  = 0
											BEGIN
												SET @psMessage = ''''
												SET @pfOK = 0
		
												RETURN
											END
										END
		
										IF @fDeletedValue = 0
										BEGIN
											SET @sSQL = ''SELECT @sValue = '' + @sTableName + ''.'' + @sColumnName +
												'' FROM '' + @sTableName +
												'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
											SET @sSQLParam = N''@sValue varchar(8000) OUTPUT''
											EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT
										END					
										IF @sValue IS null SET @sValue = ''''
							
										/* Format dates */
										IF @iSourceItemType = 11
										BEGIN
											IF len(@sValue) = 0
											BEGIN
												SET @sValue = ''<undefined>''
											END
											ELSE
											BEGIN
												SET @dtTempDate = convert(datetime, @sValue)
												SET @sValue = convert(varchar(8000), @dtTempDate, @iEmailFormat)
											END
'


	SET @sSPCode_3 = '
										END
					
										/* Format logics */
										IF @iSourceItemType = -7
										BEGIN
											IF @sValue = 0 
											BEGIN
												SET @sValue = ''False''
											END
											ELSE
											BEGIN
												SET @sValue = ''True''
											END
										END	
					
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
										SELECT @sValue = ASRSysWorkflowInstanceValues.value, 
											@iSourceItemType = ASRSysWorkflowElementItems.itemType,
											@iColumnID = ASRSysWorkflowElementItems.lookupColumnID
										FROM ASRSysWorkflowInstanceValues
										INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
										INNER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.elementID
										WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
											AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
											AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
											AND ASRSysWorkflowElementItems.identifier = @sWFValueIdentifier
							
										IF @sValue IS null SET @sValue = ''''
					
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
											WHERE ASRSysColumns.columnID = @iColumnID
										END
												
										/* Format dates */
										IF @iSourceItemType = 7
										BEGIN
											IF len(@sValue) = 0 OR @sValue = ''null''
											BEGIN
												SET @sValue = ''<undefined>''
											END
											ELSE
											BEGIN
												SET @dtTempDate = convert(datetime, @sValue)
												SET @sValue = convert(varchar(8000), @dtTempDate, @iEmailFormat)
											END
										END
							
										/* Format logics */
										IF @iSourceItemType = 6
										BEGIN
											IF @sValue = 0 
											BEGIN
												SET @sValue = ''False''
											END
											ELSE
											BEGIN
												SET @sValue = ''True''
											END
										END			
					
										SET @psMessage = @psMessage
											+ @sValue
									END
					
									IF @iItemType = 12
									BEGIN
										/* Formatting option. */
										/* NB. The empty string that precede the char codes ARE required. */
										SET @psMessage = @psMessage +
											CASE
												WHEN @sCaption = ''L'' THEN '''' + char(13) + ''--------------------------------------------------'' + char(13)
												WHEN @sCaption = ''N'' THEN '''' + char(13)
												WHEN @sCaption = ''T'' THEN '''' + char(9)
												ELSE ''''
											END
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
											0
		
										SET @psMessage = @psMessage +
											@sResult
									END
							
									FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier, @iCalcID
								END
								CLOS'


	SET @sSPCode_4 = 'E itemCursor
								DEALLOCATE itemCursor
							
								/* Append the link to the webform that follows this element (ignore connectors) if there are any. */
								CREATE TABLE #succeedingElements (elementID integer)
							
								EXEC [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]  
									@piInstanceID, 
									@piElementID, 
									@superCursor OUTPUT,
									@psTo
							
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
									SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + CHAR(13) + CHAR(13)
										+ ''Click on the following link''
										+ CASE
											WHEN @iCount = 1 THEN ''''
											ELSE ''s''
										END
										+ '' to action:''
										+ CHAR(13)
							
									DECLARE elementCursor CURSOR LOCAL FAST_FORWARD FOR 
									SELECT SE.elementID, ISNULL(WE.caption, '''')
									FROM #succeedingElements SE
									INNER JOIN ASRSysWorkflowElements WE ON SE.elementID = WE.ID
									WHERE WE.type = 2 -- 2 = Web Form element
								
									OPEN elementCursor
									FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
									WHILE (@@fetch_status = 0)
									BEGIN
			
										IF @iSQLVersion = 8
										BEGIN
											EXEC @hResult = sp_OACreate ''vbpHRProServer.clsWorkflow'', @objectToken OUTPUT
											IF @hResult <> 0
											BEGIN
												SET @sQueryString = ''''
											END
											ELSE
											BEGIN
												EXEC @hResult = sp_OAMethod @objectToken, ''GetQueryString'', @sQueryString OUTPUT, @piInstanceID, @iElementID, @sParam1, @@servername, @sDBName
												IF @hResult <> 0 
												BEGIN
													SET @sQueryString = ''''
												END
		
												EXEC @hResult = sp_OADestroy @objectToken 
											END
										END
										ELSE
										BEGIN
											SELECT @sQueryString = dbo.[udfASRNetGetWorkflowQueryString]( @piInstanceID, @iElementID, @sParam1, @@servername, @sDBName)
										END
													
										IF LEN(@sQueryString) = 0 
										BEGIN
											SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + CHAR(13) +
												@sCaption + '' - Error constructing the query string. Please contact your system administrator.''
										END
										ELSE
										BEGIN
											SET @psHypertextLinkedSteps = @psHypertextLinkedSteps
												+ CASE
													WHEN len(@psHypertextLinkedSteps) = 0 THEN char(9)
													ELSE ''''
												END 
												+ convert(varchar(8000), @iElementID)
												+ char(9)
		
											SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + CHAR(13) +
												@sCaption + '' - '' + CHAR(13) + 
												''<'' + @sURL + ''?'' + @sQueryString + ''>''
										END
										
										FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
									END
									CLOSE elementCursor
							
									DEALLOCATE elementCursor
		
									SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + CHAR(13) + CHAR(13)
										+ ''Please make sure that the link''
										+ CASE
											WHEN @iCount = 1 THEN '' has''
											ELSE ''s have''
										END
										+ '' not been cut off by your display.'' + CHAR(13)
										+ ''If ''
										+ CASE
											WHEN @iCount = 1 THEN ''it has''
											ELSE ''they have''
										END
										+ '' been cut off you will need to copy and paste ''
										+ '


	SET @sSPCode_5 = 'CASE
											WHEN @iCount = 1 THEN ''it''
											ELSE ''them''
										END
										+ '' into your browser.''
								END
							
								DROP TABLE #succeedingElements
							END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4
		+ @sSPCode_5)

	----------------------------------------------------------------------
	-- spASRGetWorkflowFormItems
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowFormItems]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowFormItems]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowFormItems]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRGetWorkflowFormItems]
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psErrorMessage	varchar(8000)	OUTPUT,
				@piBackColour	integer	OUTPUT,
				@piBackImage	integer	OUTPUT,
				@piBackImageLocation	integer	OUTPUT,
				@piWidth	integer	OUTPUT,
				@piHeight	integer	OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iID			integer,
					@iItemType		integer,
					@iDefaultValueType		integer,
					@iDBColumnID		integer,
					@iDBColumnDataType	integer,
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
					@iCount		integer,
					@iWorkflowID		integer,
					@iElementType		integer, 
					@iType integer,
					@fValidRecordID	bit,
					@iBaseTableID	integer,
					@iBaseRecordID	integer,
					@iRequiredTableID	integer,
					@iRequiredRecordID	integer,
					@iParent1TableID	int,
					@iParent1RecordID	int,
					@iParent2TableID	int,
					@iParent2RecordID	int,
					@iInitParent1TableID	int,
					@iInitParent1RecordID	int,
					@iInitParent2TableID	int,
					@iInitParent2RecordID	int,
					@fDeletedValue		bit,
					@iTempElementID		integer,
					@iColumnID	integer,
					@iResultType	integer,
					@sResult		varchar(8000),
					@fResult		bit,
					@dtResult		datetime,
					@fltResult		float,
					@iCalcID	integer,
					@iSize		integer,
					@iDecimals	integer,
					@sIdentifier	varchar(8000)

				DECLARE @itemValues table(ID integer, value varchar(8000), type integer)	
						
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
			
				SELECT 			
					@piBackColour = isnull(webFormBGColor, 16777166),
					@piBackImage = isnull(webFormBGImageID, 0),
					@piBackImageLocation = isnull(webFormBGImageLocation, 0),
					@piWidth = isnull(webFormWidth, -1),
					@piHeight = isnull(webFormHeight, -1),
					@iWorkflowID = workflowID
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
					'


	SET @sSPCode_1 = 'ASRSysWorkflowElementItems.wfValueIdentifier,
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

							SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
							FROM ASRSysModuleSetup
							WHERE moduleKey = ''MODULE_WORKFLOW''
							AND parameterKey = ''Param_TablePersonnel''
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
							'


	SET @sSPCode_2 = '	-- StoredData
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
									'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000)'


	SET @sSPCode_3 = ', @iRecordID)
							SET @sSQLParam = N''@sValue varchar(8000) OUTPUT''
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
					BEGIN
						IF @iStatus = 7 -- Previously SavedForLater
						BEGIN
							SELECT @sValue = 
								CASE
									WHEN (@iItemType = 6 AND IVs.value = ''1'') THEN ''TRUE'' 
									WHEN (@iItemType = 6 AND IVs.value <> ''1'') THEN ''FALSE'' 
									WHEN (@iItemType = 7 AND (upper(ltrim(rtrim(IVs.value))) = ''NULL'')) THEN '''' 
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
										WHEN @iResultType = 2 THEN STR(@fltResult, 8000, @iDecimals)
										WHEN @iResultType = 3 THEN 
											CASE 
												WHEN @fResult = 1 THEN ''TRUE''
												ELSE ''FALSE''
											END
										WHEN @iResultType = 4 THEN convert(varchar(100), @dtResult, 101)
										ELSE convert(varchar(8000), @sResult)
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
					'


	SET @sSPCode_4 = '	@iItemType, 
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
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4)

	----------------------------------------------------------------------
	-- spASRGetWorkflowGridItems
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowGridItems]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowGridItems]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowGridItems]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetWorkflowGridItems
			(
				@piInstanceID		integer,
				@piElementItemID	integer, 
				@pfOK				bit	OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iTableID 		integer,
					@iOrderID			integer,
					@iFilterID			integer,
					@sFilterSQL			varchar(8000),
					@sFilterUDF			varchar(8000),
					@sRecSelWebFormIdentifier	varchar(200),
					@sRecSelIdentifier	varchar(200),
					@iDBRecord		integer,
					@iInitiatorID		integer,
					@sSQL			varchar(8000),
					@sOrderItemType	varchar(8000),
					@sSelectSQL		varchar(8000),
					@sOrderSQL		varchar(8000),
					@sBaseTableName	sysname,
					@fAscending		bit,
					@sColumnName		sysname,
					@sTempTableName	sysname,
					@iTempTableID		integer,
					@iTempTableType	integer,
					@iTempCount		integer,
					@sValue			varchar(8000),
					@iDataType		integer,
					@iRecordID		integer,
					@iPersonnelTableID	integer,
					@iWorkflowID		integer,
					@iElementType		integer, 
					@fValidRecordID	bit,
					@iElementID	integer,
					@iBaseTableID	integer,
					@iParent1TableID	int,
					@iParent1RecordID	int,
					@iParent2TableID	int,
					@iParent2RecordID	int,
					@iRecordTableID		int,
					@iTriggerTableID	integer
				DECLARE @joinParents table(tableID		integer)	
			
				SET @pfOK = 1

				SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel''
			
				SELECT 			
					@iTableID = ASRSysWorkflowElementItems.tableID,
					@iElementID = ASRSysWorkflowElementItems.elementiD,
					@sRecSelWebFormIdentifier = isnull(ASRSysWorkflowElementItems.wfFormIdentifier, ''''),
					@sRecSelIdentifier = isnull(ASRSysWorkflowElementItems.wfValueIdentifier, 0),
					@iDBRecord = ASRSysWorkflowElementItems.dbRecord,
					@iOrderID = 
						CASE
							WHEN isnull(ASRSysWorkflowElementItems.recordOrderID, 0) > 0 THEN ASRSysWorkflowElementItems.recordOrderID
							ELSE ASRSysTables.defaultOrderID
						END,
					@iFilterID = isnull(ASRSysWorkflowElementItems.recordFilterID, 0),
					@iRecordTableID = ASRSysWorkflowElementItems.recordTableID,
					@sBaseTableName = ASRSysTables.tableName
				FROM ASRSysWorkflowElementItems
				INNER JOIN ASRSysTables ON ASRSysWorkflowElementItems.tableID = ASRSysTables.tableID
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iWorkflowID = ASRSysWorkflowInstances.workflowID, 
					@iTriggerTableID = ASRSysWorkflows.baseTable,
					@iParent1TableID = ASRSysWorkflowInstances.parent1TableID,
					@iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
					@iParent2TableID = ASRSysWorkflowInstances.parent2TableID,
					@iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
				FROM ASRSysWorkflowInstances
				INNER JOIN ASRSysWorkflows ON ASRSysWorkflowInstances.workflowID = ASRSysWorkflows.ID
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				SET @sSelectSQL = ''''
				SET @sOrderSQL = ''''
			
				DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT 
					ASRSysColumns.columnName,
					ASRSysColumns.dataType,
					ASRSysColumns.tableID,
					ASRSysTables.tableType,
					ASRSysTables.tableName,
					upper(isnull(ASRSysOrderItems.type, '''')),
					ASRSysOrderItems.ascending
				FROM ASRSysOrderItems
				INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
				INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
				WHERE ASRSysOrderItems.orderID = @iOrderID
				ORDER BY ASRSysOrderItems.type,
					ASRSysOrderItems.sequence
			
				OPEN orderCursor
				FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTempTableID, @iTempTableType, @sTempTableName, @sOrderItemType, @fAscending
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @sOrderItemType = ''F''
'


	SET @sSPCode_1 = '					BEGIN
						SET @sSelectSQL = @sSelectSQL +
							CASE 
								WHEN len(@sSelectSQL) > 0 THEN '',''
								ELSE ''''
							END +
							@sTempTableName + ''.'' + @sColumnName
					END
			
					IF @sOrderItemType = ''O''
					BEGIN
						SET @sOrderSQL = @sOrderSQL + 
							CASE 
								WHEN len(@sOrderSQL) > 0 THEN '',''
								ELSE '' ''
							END + 
							@sTempTableName + ''.'' + @sColumnName +
							CASE 
								WHEN @fAscending = 0 THEN '' DESC'' 
								ELSE '''' 
							END				
					END
			
					IF @iTableID <> @iTempTableID
					BEGIN
						SELECT @iTempCount = COUNT(tableID)
						FROM @joinParents
						WHERE tableID = @iTempTableID
			
						IF @iTempCount = 0
						BEGIN
							INSERT INTO @joinParents (tableID) VALUES(@iTempTableID)
						END
					END
			
					FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTempTableID, @iTempTableType, @sTempTableName, @sOrderItemType, @fAscending
				END
				CLOSE orderCursor
				DEALLOCATE orderCursor
			
				IF len(@sSelectSQL) > 0 
				BEGIN
					SET @sSelectSQL = ''SELECT '' + @sSelectSQL + '','' +
						@sBaseTableName + ''.id'' +
					'' FROM '' + @sBaseTableName
			
					DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT ASRSysTables.tableName, 
						JP.tableID
					FROM @joinParents JP
					INNER JOIN ASRSysTables ON JP.tableID = ASRSysTables.tableID
			
					OPEN joinCursor
					FETCH NEXT FROM joinCursor INTO @sTempTableName, @iTempTableID
					WHILE (@@fetch_status = 0)
					BEGIN
						SET @sSelectSQL = @sSelectSQL + 
							'' LEFT OUTER JOIN '' + @sTempTableName + '' ON '' + @sBaseTableName + ''.ID_'' + convert(varchar(100), @iTempTableID) + '' = '' + @sTempTableName + ''.ID''
			
						FETCH NEXT FROM joinCursor INTO @sTempTableName, @iTempTableID
					END
					CLOSE joinCursor
					DEALLOCATE joinCursor
			
					IF @iDBRecord = 0 -- ie. based on the initiator''s record
					BEGIN
						SET @iBaseTableID = @iPersonnelTableID
						SET @iRecordID = @iInitiatorID
					END

					IF @iDBRecord = 4 -- ie. based on the triggered record
					BEGIN
						SET @iBaseTableID = @iTriggerTableID
						SET @iRecordID = @iInitiatorID
					END
			
					IF @iDBRecord = 1 -- ie. based on a previously identified record
					BEGIN
						SELECT @iElementType = ASRSysWorkflowElements.type
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))
		
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
								AND IV.elementID = Es.ID
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
							WHERE IV.instanceID = @piInstance'


	SET @sSPCode_2 = 'ID
						END
			
						SET @iRecordID = 
							CASE
								WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
								ELSE 0
							END

						SET @iBaseTableID = @iTempTableID
					END
			
					IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
					BEGIN
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
							@iBaseTableID,
							@iRecordID,
							@iParent1TableID,
							@iParent1RecordID,
							@iParent2TableID,
							@iParent2RecordID,
							@iRecordTableID,
							@iRecordID	OUTPUT

						SET @sSelectSQL = @sSelectSQL + 
							'' WHERE '' + @sBaseTableName + ''.ID_'' + convert(varchar(100), @iRecordTableID) + '' = '' + convert(varchar(100), @iRecordID)

						SET @fValidRecordID = 1

						EXEC [dbo].[spASRWorkflowValidTableRecord]
							@iRecordTableID,
							@iRecordID,
							@fValidRecordID	OUTPUT

						IF @fValidRecordID  = 0
						BEGIN
							SET @pfOK = 0

							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC [dbo].[spASRWorkflowActionFailed] @piInstanceID, @iElementID, ''Web Form record selector item record has been deleted or not selected.''
							
							-- Need to return a recordset of some kind.
							SELECT '''' AS ''Error''

							RETURN
						END
					END

					IF @iFilterID > 0 
					BEGIN
						SET @sFilterUDF = ''[dbo].udf_ASRWFExpr_'' + convert(varchar(8000), @iFilterID)

						IF EXISTS(
							SELECT Name
							FROM sysobjects
							WHERE id = object_id(@sFilterUDF)
							AND sysstat & 0xf = 0)
						BEGIN
							SET @sFilterSQL = 
								CASE
									WHEN (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4) THEN '' AND ''
									ELSE '' WHERE ''
								END 
								+ @sBaseTableName + ''.ID  IN (SELECT id FROM '' + @sFilterUDF + ''('' + convert(varchar(8000), @piInstanceID) + ''))''
						END
					END

					SET @sOrderSQL = '' ORDER BY '' + @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN '','' 
							ELSE '''' 
						END + 
						@sBaseTableName + ''.ID''

					EXEC (@sSelectSQL 
						+ @sFilterSQL
						+ @sOrderSQL)
				END
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

	----------------------------------------------------------------------
	-- spASRSubmitWorkflowStep
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSubmitWorkflowStep]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSubmitWorkflowStep]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRSubmitWorkflowStep]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRSubmitWorkflowStep
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@psFormInput1		varchar(8000),
			@psFormInput2		varchar(8000),
			@psFormElements		varchar(8000)	OUTPUT,
			@pfSavedForLater	bit				OUTPUT
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
				@hTmpResult		integer,
				@sTo			varchar(8000),
				@sCopyTo		varchar(8000),
				@sTempTo		varchar(8000),
				@sMessage		varchar(8000),
				@sMessage_HypertextLinks	varchar(8000),
				@sHypertextLinkedSteps		varchar(8000),
				@iEmailID		integer,
				@iEmailCopyID		integer,
				@iTempEmailID		integer,
				@iEmailLoop		integer,
				@iEmailRecord		integer,
				@iEmailRecordID	integer,
				@sSQL			nvarchar(4000),
				@iCount		integer,
				@superCursor		cursor,
				@curDelegatedRecords	cursor,
				@fDelegate		bit,
				@fDelegationValid	bit,
				@iDelegateEmailID	integer,
				@iDelegateRecordID	integer,
				@sTemp		varchar(8000),
				@sDelegateTo		varchar(8000),
				@sAllDelegateTo	varchar(8000),
				@iCurrentStepID	int,
				@sDelegatedMessage	varchar(8000),
				@iTemp		integer, 
				@iPrevElementType	integer,
				@iWorkflowID		integer,
				@sRecSelIdentifier	varchar(8000),
				@sRecSelWebFormIdentifier	varchar(8000), 
				@iStepID int,
				@iElementID int,
				@sUserName varchar(8000),
				@sUserEmail varchar(8000), 
				@sValueDescription	varchar(8000),
				@iTableID		integer,
				@iRecDescID		integer,
				@sEvalRecDesc	varchar(8000),
				@sExecString		nvarchar(4000),
				@sParamDefinition	nvarchar(500),
				@sIdentifier		varchar(8000),
				@iItemType		integer,
				@iDataAction		integer, 
				@fValidRecordID	bit,
				@iEmailTableID integer,
				@iEmailType integer,
				@iBaseTableID	integer,
				@iBaseRecordID	integer,
				@iRequiredRecordID	integer,
				@iParent1TableID	int,
				@iParent1RecordID	int,
				@iParent2TableID	int,
				@iParent2RecordID	int,
				@iTempElementID		integer,
				@iTrueFlowType	integer,
				@iExprID		integer,
				@iResultType	integer,
				@sResult		varchar(8000),
				@fResult		bit,
				@dtResult		datetime,
				@fltResult		float,
				@sEmailSubject	varchar(200),
				@iTempID	integer,
				@iBehaviour		integer

			SET @pfSavedForLater = 0

			SELECT @iCurrentStepID = ID
			FROM ASRSysWorkflowInstanceSteps
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

			SET @iDelegateEmailID = 0
			SELECT @sTemp = ISNULL(parameterValue, '''')
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_DelegateEmail''
			SET @iDelegateEmailID = convert(integer, @sTemp)

			SET @psFormElements = ''''
						
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
			WHERE E.ID = @piElementID

			--------------------------------------------------
			-- Read the submitted webForm/storedData values
			--------------------------------------------------
			IF @iElementType = 5 -- Stored Data element
			BEGIN
				SET @sValue = @psFormInput1
				SET @sValueDescription = ''''
				SET @sMessage = ''Successfully '' +
					CASE
						WHEN @iDataAction = 0 THEN ''inserted''
						WHEN @iDataAction = 1 THEN '


	SET @sSPCode_1 = '''updated''
						ELSE ''deleted''
					END + '' record''

				IF @iDataAction = 2 -- Deleted - Record Description calculated before the record was deleted.
				BEGIN
					SET @sValueDescription = @psFormInput2
				END
				ELSE
				BEGIN
					SET @iTemp = convert(integer, @sValue)
					IF @iTemp > 0 
					BEGIN	
						EXEC [dbo].[spASRRecordDescription] 
							@iTableID,
							@iTemp,
							@sEvalRecDesc OUTPUT
						IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc
					END
				END

				IF len(@sValueDescription) > 0 SET @sMessage = @sMessage + '' ('' + @sValueDescription + '')''

				UPDATE ASRSysWorkflowInstanceValues
				SET ASRSysWorkflowInstanceValues.value = @sValue, 
					ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription
				WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceValues.elementID = @piElementID
					AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
					AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0
			END
			ELSE
			BEGIN
				-- Put the submitted form values into the ASRSysWorkflowInstanceValues table. 
				WHILE (charindex(CHAR(9), @psFormInput1) > 0) OR (charindex(CHAR(9), @psFormInput2) > 0)
				BEGIN
					SET @iIndex1 = charindex(CHAR(9), @psFormInput1)
					IF @iIndex1 > 0
					BEGIN
						SET @sID = replace(LEFT(@psFormInput1, @iIndex1-1), '''''''', '''''''''''')

						SET @iIndex2 = charindex(CHAR(9), @psFormInput1, @iIndex1+1)
						IF @iIndex2 > 0	
						BEGIN
							SET @sValue = SUBSTRING(@psFormInput1, @iIndex1+1, @iIndex2-@iIndex1-1)

							SET @psFormInput1 = SUBSTRING(@psFormInput1, @iIndex2+1, LEN(@psFormInput1) - @iIndex2)
						END
						ELSE
						BEGIN
							SET @iIndex2 = charindex(CHAR(9), @psFormInput2)
							SET @sValue = SUBSTRING(@psFormInput1, @iIndex1+1, len(@psFormInput1)-@iIndex1) +
								LEFT(@psFormInput2, @iIndex2-1)

							SET @psFormInput1 = ''''
							SET @psFormInput2 = SUBSTRING(@psFormInput2, @iIndex2+1, LEN(@psFormInput2) - @iIndex2)
						END
					END
					ELSE
					BEGIN
						SET @iIndex1 = charindex(CHAR(9), @psFormInput2)
						SET @iIndex2 = charindex(CHAR(9), @psFormInput2, @iIndex1+1)

						SET @sID = replace(@psFormInput1, '''''''', '''''''''''') +
							replace(LEFT(@psFormInput2, @iIndex1-1), '''''''', '''''''''''')
						SET @sValue = SUBSTRING(@psFormInput2, @iIndex1+1, @iIndex2-@iIndex1-1)

						SET @psFormInput1 = ''''
						SET @psFormInput2 = SUBSTRING(@psFormInput2, @iIndex2+1, LEN(@psFormInput2) - @iIndex2)
					END
					SET @sValue = left(@sValue, 8000)

					--Get the record description (for RecordSelectors only)
					SET @sValueDescription = ''''

					-- Get the WebForm item type, etc.
					SELECT @sIdentifier = EI.identifier,
						@iItemType = EI.itemType,
						@iTableID = EI.tableID,
						@iBehaviour = EI.behaviour
					FROM ASRSysWorkflowElementItems EI
					WHERE EI.ID = convert(integer, @sID)

					SET @iParent1TableID = 0
					SET @iParent1RecordID = 0
					SET @iParent2TableID = 0
					SET @iParent2RecordID = 0

					IF @iItemType = 11 -- Record Selector
					BEGIN
						-- Get the table record description ID. 
						SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
						FROM ASRSysTables 
						WHERE ASRSysTables.tableID = @iTableID

						SET @iTemp = convert(integer, isnull(@sValue, ''0''))

						-- Get the record description. 
						IF (NOT @iRecDescID IS null) AND (@iRecDescID > 0) AND (@iTemp > 0)
						BEGIN
							SET @sExecString = ''exec sp_ASRExpr_'' + convert(nvarchar(4000), @iRecDescID) + '' @recDesc OUTPUT, @recID''
							SET @sParamDefinition = N''@recDesc varchar(8000) OUTPUT, @recID integer''
							EXEC sp_executesql @sExecString, @sParamDefinition, @sEvalRecDesc OUTPUT, @iTemp
							IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalR'


	SET @sSPCode_2 = 'ecDesc
						END

						-- Record the selected record''s parent details.
						exec [dbo].[spASRGetParentDetails]
							@iTableID,
							@iTemp,
							@iParent1TableID	OUTPUT,
							@iParent1RecordID	OUTPUT,
							@iParent2TableID	OUTPUT,
							@iParent2RecordID	OUTPUT
					END
					ELSE
					IF (@iItemType = 0) and (@iBehaviour = 1) AND (@sValue = ''1'')-- SaveForLater Button
					BEGIN
						SET @pfSavedForLater = 1
					END

					UPDATE ASRSysWorkflowInstanceValues
					SET ASRSysWorkflowInstanceValues.value = @sValue, 
						ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription,
						ASRSysWorkflowInstanceValues.parent1TableID = @iParent1TableID,
						ASRSysWorkflowInstanceValues.parent1RecordID = @iParent1RecordID,
						ASRSysWorkflowInstanceValues.parent2TableID = @iParent2TableID,
						ASRSysWorkflowInstanceValues.parent2RecordID = @iParent2RecordID
					WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceValues.elementID = @piElementID
						AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier
				END

				IF @pfSavedForLater = 1
				BEGIN
					/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 7
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

					RETURN
				END
			END
					
			SET @hResult = 0
			SET @sTo = ''''
			SET @sCopyTo = ''''
		
			--------------------------------------------------
			-- Process email element
			--------------------------------------------------
			IF @iElementType = 3 -- Email element
			BEGIN
				-- Get the email recipient. 
				SET @iEmailRecordID = 0
				SET @sSQL = ''spASRSysEmailAddr''

				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @iEmailLoop = 0
					WHILE @iEmailLoop < 2
					BEGIN
						SET @hTmpResult = 0
						SET @sTempTo = ''''
						SET @iTempEmailID = 
							CASE 
								WHEN @iEmailLoop = 1 THEN @iEmailCopyID
								ELSE isnull(@iEmailID, 0)
							END

						IF @iTempEmailID > 0 
						BEGIN
							SET @fValidRecordID = 1

							SELECT @iEmailTableID = isnull(tableID, 0),
								@iEmailType = isnull(type, 0)
							FROM ASRSysEmailAddress
							WHERE emailID = @iTempEmailID

							IF @iEmailType = 0 
							BEGIN
								SET @iEmailRecordID = 0
							END
							ELSE
							BEGIN
								SET @iTempElementID = 0

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
									WHERE ASRSysWorkflowInstances.ID = @piInstanceID

									SET @iBaseRecordID = @iEmailRecordID

									IF @iEmailRecord = 4
									BEGIN
										-- Trigger record
										SELECT @iBaseTableID = isnull(WF.baseTable, 0)
										FROM ASRSysWorkflows WF
										INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
											AND WFI.ID = @piInstanceID
									END
									ELSE
									BEGIN
										-- Initiator''s record
										SELECT @iBaseTableID = convert(integer, isnull(parameterValue, 0))
										FROM ASRSysModuleSetup
										WHERE moduleKey = ''MODULE_WORKFLOW''
										AND parameterKey = ''Param_TablePersonnel''
									END
								END
		
								IF @iEmailRecord = 1
								BEGIN
									SELECT @iPrevElementType'


	SET @sSPCode_3 = ' = ASRSysWorkflowElements.type,
										@iTempElementID = ASRSysWorkflowElements.ID
									FROM ASRSysWorkflowElements
									WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
										AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))

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
											AND Es.identifier = @sRecSelWebFormIdentifier
										WHERE IV.instanceID = @piInstanceID
									END

									SET @iEmailRecordID = 
										CASE
											WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
											ELSE 0
										END

									SET @iBaseRecordID = @iEmailRecordID
								END

								SET @fValidRecordID = 1
								IF (@iEmailRecord = 0) OR (@iEmailRecord = 1) OR (@iEmailRecord = 4)
								BEGIN
									SET @fValidRecordID = 0

									EXEC [dbo].[spASRWorkflowAscendantRecordID]
										@iBaseTableID,
										@iBaseRecordID,
										@iParent1TableID,
										@iParent1RecordID,
										@iParent2TableID,
										@iParent2RecordID,
										@iEmailTableID,
										@iRequiredRecordID	OUTPUT

									SET @iEmailRecordID = @iRequiredRecordID

									IF @iRequiredRecordID > 0 
									BEGIN
										EXEC [dbo].[spASRWorkflowValidTableRecord]
											@iEmailTableID,
											@iEmailRecordID,
											@fValidRecordID	OUTPUT
									END

									IF @fValidRecordID = 0
									BEGIN
										IF @iEmailRecord = 4 -- Trigger record. See if the email address was calulated as part of the delete trigger.
										BEGIN
											SELECT @sTempTo = rtrim(ltrim(isnull(QC.columnValue , '''')))
											FROM ASRSysWorkflowQueueColumns QC
											INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
											WHERE WFQ.instanceID = @piInstanceID
												AND QC.emailID = @iTempEmailID

											IF len(@sTempTo) > 0 SET @fValidRecordID = 1
										END
										ELSE
										BEGIN
											IF @iEmailRecord = 1
											BEGIN
												SELECT @sTempTo = rtrim(ltrim(isnull(IV.value , '''')))
												FROM ASRSysWorkflowInstanceValues IV
												WHERE IV.instanceID = @piInstanceID
													AND IV.emailID = @iTempEmailID
													AND IV.elementID = @iTempElementID

												IF len(@sTempTo) > 0 SET @fValidRecordID = 1
											END
										END
									END

									IF (@fValidRecordID = 0) AND (@iEmailLoop = 0)
									BEGIN
										-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
										EXEC [dbo].[spASRWorkflowActionFailed] 
	'


	SET @sSPCode_4 = '										@piInstanceID, 
											@piElementID, 
											''Email record has been deleted or not selected.''
													
										SET @hTmpResult = -1
									END
								END
							END

							IF @fValidRecordID = 1
							BEGIN
								/* Get the recipient address. */
								IF len(@sTempTo) = 0
								BEGIN
									EXEC @hTmpResult = @sSQL @sTempTo OUTPUT, @iTempEmailID, @iEmailRecordID
									IF @sTempTo IS null SET @sTempTo = ''''
								END

								IF (LEN(rtrim(ltrim(@sTempTo))) = 0) AND (@iEmailLoop = 0)
								BEGIN
									-- Email step failure if no known recipient.
									-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
									EXEC [dbo].[spASRWorkflowActionFailed] 
										@piInstanceID, 
										@piElementID, 
										''No email recipient.''
												
									SET @hTmpResult = -1
								END
							END

							IF @iEmailLoop = 1 
							BEGIN
								SET @sCopyTo = @sTempTo

								IF (rtrim(ltrim(@sCopyTo)) = ''@'')
									OR (charindex('' @ '', @sCopyTo) > 0)
								BEGIN
									SET @sCopyTo = ''''
								END
							END
							ELSE
							BEGIN
								SET @sTo = @sTempTo
							END
						END
						
						SET @iEmailLoop = @iEmailLoop + 1

						IF @hTmpResult <> 0 SET @hResult = @hTmpResult
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
							AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

						EXEC [dbo].[spASRWorkflowActionFailed] 
							@piInstanceID, 
							@piElementID, 
							''Invalid email recipient.''
						
						SET @hResult = -1
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
							@sTo
		
						IF @fValidRecordID = 1
						BEGIN
							exec [dbo].[spASRDelegateWorkflowEmail] 
								@sTo,
								@sCopyTo,
								@sMessage,
								@sMessage_HypertextLinks,
								@iCurrentStepID,
								@sEmailSubject
						END
						ELSE
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC [dbo].[spASRWorkflowActionFailed] 
								@piInstanceID, 
								@piElementID, 
								''Email item database value record has been deleted or not selected.''
										
							SET @hResult = -1
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
						WHEN @iElementType = 3 THEN LEFT(@sMessage, 8000)
						WHEN @iElementType = 5'


	SET @sSPCode_5 = ' THEN LEFT(@sMessage, 8000)
						ELSE ''''
					END,
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
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
							0

						SET @iValue = convert(integer, @fResult)
					END
					ELSE
					BEGIN
						-- Decision Element flow determined by a button in a preceding web form
						SET @iPrevElementType = 4 -- Decision element
						SET @iPreviousElementID = @piElementID

						WHILE (@iPrevElementType = 4)
						BEGIN
							SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
								@iPrevElementType = isnull(WE.type, 0)
							FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPreviousElementID) PE
							INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
							INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
								AND WIS.instanceID = @piInstanceID

							SET @iPreviousElementID = @iTempID
						END
					
						SELECT @sValue = ISNULL(IV.value, ''0'')
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
						WHERE IV.elementID = @iPreviousElementID
							AND IV.instanceid = @piInstanceID
							AND E.ID = @piElementID

						SET @iValue = 
							CASE
								WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
								ELSE 0
							END
					END
				
					IF @iValue IS null SET @iValue = 0
		
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
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
							OR ASRSysWorkflowInstanceSteps.status = 3)
				END
				ELSE
				BEGIN
					IF @iElementType <> 3 -- 3=Email element
					BEGIN
						-- Do not the following bit when the submitted element is an Email element as 
						-- the succeeding elements will already have been actioned.
						DECLARE @succeedingElements TABLE(elementID integer)
		
						EXEC [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]  
							@piInstanceID, 
							@piElementID, 
							@superCursor OUTPUT,
							''''
		
						FETCH NEXT FROM @superCursor INTO @iTemp
						WHILE (@@fetch_status = 0)
						BEGIN
							INSERT INTO @succeedingElements (elementID) VALUES (@iTemp)
							
							FETCH NEXT FROM @superCursor INTO @iTemp 
						END
						CLOSE @superCursor
						DEALLOCATE @superCursor

						-- If the submitted element is a web form, then any succeeding webforms are actioned for the same user.
						IF @iElementType = 2 -- WebForm
						BEGIN
							SELECT @sUserName = isnull(WIS.userName, ''''),
								@sUserEmail = isnull(WIS.userEmail, '''')
							FROM ASRSysWorkflowInstanceSteps WIS
							WHERE WIS.instanceID = @piInstanceID
								AND WIS.elementI'


	SET @sSPCode_6 = 'D = @piElementID

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
									OR ASRSysWorkflowInstanceSteps.status = 3)

							OPEN formsCursor
							FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
							WHILE (@@fetch_status = 0) 
							BEGIN
								SET @psFormElements = @psFormElements + convert(varchar(8000), @iElementID) + char(9)

								DELETE FROM ASRSysWorkflowStepDelegation
								WHERE stepID = @iStepID

								INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
									(SELECT WSD.delegateEmail, @iStepID
									FROM ASRSysWorkflowStepDelegation WSD
									WHERE WSD.stepID = @iCurrentStepID)
								
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
										OR ASRSysWorkflowInstanceSteps.status = 3)
								
								FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
							END
							CLOSE formsCursor
							DEALLOCATE formsCursor

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
									OR ASRSysWorkflowInstanceSteps.status = 3)
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
										OR ASRSysWorkflowInstanceSteps.status = 3))
							
							INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
	'


	SET @sSPCode_7 = '						(SELECT WSD.delegateEmail,
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
							WHERE WSD.stepID = @iCurrentStepID)

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
									OR ASRSysWorkflowInstanceSteps.status = 3)
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
						AND ASRSysWorkflowElements.type = 2)
		
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
						AND ASRSysWorkflowElements.type = 1)
		
				-- Count how many terminators have completed. ie. if the workflow has completed. 
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type = 1
							
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @piInstanceID
					
					-- Steps pending action are no longer required.
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 0 -- 0 = On hold
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND (ASRSysWorkflowInstanceSteps.status = 1 -- 1 = Pending Engine Action
							OR ASRSysWorkflowInstanceSteps.status = '


	SET @sSPCode_8 = '2) -- 2 = Pending User Action
				END

				IF @iElementType = 3 -- Email element
					OR @iElementType = 5 -- Stored Data element
				BEGIN
					exec [dbo].[spASREmailImmediate] ''HR Pro Workflow''
				END
			END
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4
		+ @sSPCode_5
		+ @sSPCode_6
		+ @sSPCode_7
		+ @sSPCode_8)

	----------------------------------------------------------------------
	-- spASRWorkflowSubmitImmediatesAndGetSucceedingElements
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@succeedingElements	cursor varying output,
			@psTo				varchar(8000)
		)
		AS
		BEGIN
			-- Action any immediate elements (Ors & Decisions) and return the IDs of the workflow elements that 
			-- succeed them.
			-- This ignores connection elements.
			DECLARE
				@iTempID		integer,
				@iElementID		integer,
				@iElementType	integer,
				@iFlowCode		integer,
				@iTrueFlowType	integer,
				@iExprID		integer,
				@iResultType	integer,
				@sValue			varchar(8000),
				@sResult		varchar(8000),
				@fResult		bit,
				@dtResult		datetime,
				@fltResult		float,
				@iValue			integer,
				@iPrecedingElementType	integer, 
				@iPrecedingElementID	integer, 
				@iCount			integer,
				@iStepID		integer,
				@curRecipients		cursor,
				@sEmailAddress		varchar(8000),
				@fDelegated			bit,
				@sDelegatedTo		varchar(8000),
				@fIsDelegate		bit
		
			DECLARE @elements table
			(
				elementID		integer,
				elementType		integer,
				processed		tinyint default 0,
				trueFlowType	integer,
				trueFlowExprID	integer
			)
							
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
			INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
				
			SELECT @iCount = COUNT(*)
			FROM @elements
			WHERE (elementType = 4 OR elementType = 7) -- 4=Decision, 7=Or
				AND processed = 0
		
			WHILE @iCount > 0
			BEGIN
				UPDATE @elements
				SET processed = 1
				WHERE processed = 0
		
				-- Action any succeeding immediate elements (Decisions & Ors)
				DECLARE immediateCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.elementID,
					E.elementType,
					E.trueFlowType, 
					E.trueFlowExprID
				FROM @elements E
				WHERE (E.elementType = 4 OR E.elementType = 7) -- 4=Decision, 7=Or
					AND E.processed = 1
		
				OPEN immediateCursor
				FETCH NEXT FROM immediateCursor INTO 
					@iElementID, 
					@iElementType, 
					@iTrueFlowType, 
					@iExprID
				WHILE (@@fetch_status = 0)
				BEGIN
					-- Submit the immediate elements, and get their succeeding elements
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 3,
						ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
						ASRSysWorkflowInstanceSteps.userEmail = ASRSysWorkflowInstanceSteps.userEmail,
						ASRSysWorkflowInstanceSteps.message = '''',
						ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @iElementID
		
					SET @iFlowCode = 0
		
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
								0
		
							SET @iValue = convert(integer, @fResult)
						END
						ELSE
						BEGIN
							-- Decision Element flow determined by a button in a preceding web form
							SET @iPrecedingElementType = 4 -- Decision element
							SET @iPrecedingElementID = @iElementID
		
							WHILE (@iPrecedingElementType = 4)
							BEGIN
								SELECT TOP 1 @iTempID = isnull(WE.ID, 0),
									@iPrecedingElementType = isnull(WE.type, 0)
								FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iPrecedingElementID) PE
								INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
								INNER'


	SET @sSPCode_1 = ' JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
									AND WIS.instanceID = @piInstanceID
		
								SET @iPrecedingElementID = @iTempID
							END
							
							SELECT @sValue = ISNULL(IV.value, ''0'')
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
							WHERE IV.elementID = @iPrecedingElementID
							AND IV.instanceid = @piInstanceID
								AND E.ID = @iElementID
		
							SET @iValue = 
								CASE
									WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
									ELSE 0
								END
						END
						
						IF @iValue IS null SET @iValue = 0
						SET @iFlowCode = @iValue
		
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @iElementID
					END
					ELSE IF @iElementType = 7 -- Or
					BEGIN
						EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @piInstanceID, @iElementID
					END
		
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
					WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements)
		
					FETCH NEXT FROM immediateCursor INTO 
						@iElementID, 
						@iElementType, 
						@iTrueFlowType, 
						@iExprID
				END
				CLOSE immediateCursor
				DEALLOCATE immediateCursor
		
				UPDATE @elements
				SET processed = 2
				WHERE processed = 1
		
				SELECT @iCount = COUNT(*)
				FROM @elements
				WHERE (elementType = 4 OR elementType = 7) -- 4=Decision, 7=Or
					AND processed = 0
			END
		
			SELECT @iCount = COUNT(*)
			FROM @elements
			WHERE elementType = 2 -- 2=WebForm
		
			IF (@iCount > 0) AND len(ltrim(rtrim(@psTo))) > 0 
			BEGIN
				SELECT @iStepID = ASRSysWorkflowInstanceSteps.ID
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
		
				DECLARE @recipients TABLE (
					emailAddress	varchar(8000),
					delegated		bit,
					delegatedTo		varchar(8000),
					isDelegate		bit
				)
		
				exec [dbo].[spASRGetWorkflowDelegates] 
					@psTo, 
					@iStepID, 
					@curRecipients output
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
					)
					
					FETCH NEXT FROM @curRecipients INTO 
							@sEmailAddress,
							@fDelegated,
							@sDelegatedTo,
							@fIsDelegate
				END
				CLOSE @curRecipients
				DEALLOCATE @curRecipients
		
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
							OR ASRSysWorkflowInstanceSteps.status = 3))
		
				INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
				S'


	SET @sSPCode_2 = 'ELECT DISTINCT RECS.emailAddress, WIS.ID
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
							OR WIS.status = 3)
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
						AND E.elementType <> 4) -- 4 = Decision
				AND (ASRSysWorkflowInstanceSteps.status = 0
					OR ASRSysWorkflowInstanceSteps.status = 2
					OR ASRSysWorkflowInstanceSteps.status = 6
					OR ASRSysWorkflowInstanceSteps.status = 8
					OR ASRSysWorkflowInstanceSteps.status = 3)
		
			UPDATE ASRSysWorkflowInstanceSteps
			SET ASRSysWorkflowInstanceSteps.status = 2
			WHERE ASRSysWorkflowInstanceSteps.id IN (
				SELECT ASRSysWorkflowInstanceSteps.ID
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND ASRSysWorkflowElements.type = 2)
		
			-- Return the cursor of succeeding elements. 
			SET @succeedingElements = CURSOR FORWARD_ONLY STATIC FOR
				SELECT elementID 
				FROM @elements E
				WHERE E.elementType <> 7 -- 7 = Or
					AND E.elementType <> 4 -- 4 = Decision
		
			OPEN @succeedingElements
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

	----------------------------------------------------------------------
	-- spASRActionActiveWorkflowSteps
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRActionActiveWorkflowSteps]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
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
				@sMessage			varchar(8000),
				@iTemp				integer, 
				@iTemp2				integer, 
				@iTemp3				integer,
				@sForms 			varchar(8000), 
				@iType				integer,
				@iDecisionFlow		integer,
				@fInvalidElements	bit, 
				@fValidElements	bit, 
				@iPrecedingElementID	integer, 
				@iPrecedingElementType	integer, 
				@iPrecedingElementStatus	integer, 
				@iPrecedingElementFlow	integer, 
				@fSaveForLater		bit
		
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
					-- Check if all preceding steps have completed before submitting this step.
					SET @fInvalidElements = 0		
				
					DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT WE.ID,
						WE.type,
						WIS.status,
						WIS.decisionFlow
					FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
					INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
					INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
						AND WIS.instanceID = @iInstanceID

					OPEN precedingElementsCursor
			
					FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow

					WHILE (@@fetch_status = 0)
						AND (@fInvalidElements = 0)
					BEGIN
						IF (@iPrecedingElementType = 4) -- Decision
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
							BEGIN
								SELECT @iCount = COUNT(*) 
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
								WHERE ID = @iElementID

								IF @iCount = 0 SET @fInvalidElements = 1
							END
							ELSE
							BEGIN
								SET @fInvalidElements = 1
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
										END

									SELECT @iCount = COUNT(*)
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
									WHERE ID = @iElementID
								
									IF @iCount = 0 SET @fInvalidElements = 1
								END
								ELSE
								BEGIN
									SET @fInvalidElements = 1
								END
							END
							ELSE
							BEGIN
								IF (@iPrecedingElementType = 5) -- StoredData
								BEGIN
									IF @iPrecedingE'


	SET @sSPCode_1 = 'lementStatus = 3 -- 3 = completed
										OR @iPrecedingElementStatus = 8 -- 8 = failed action
									BEGIN
										SET @iTemp3 = CASE
												WHEN @iPrecedingElementStatus = 3 THEN 0
												ELSE 1
											END

										SELECT @iCount = COUNT(*)
										FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
										WHERE ID = @iElementID
									
										IF @iCount = 0 SET @fInvalidElements = 1
									END
									ELSE
									BEGIN
										SET @fInvalidElements = 1
									END
								END
								ELSE
								BEGIN
									-- Preceding element must have status 3 (3 =Completed)
									IF @iPrecedingElementStatus <> 3 SET @fInvalidElements = 1
								END
							END
						END

						FETCH NEXT FROM precedingElementsCursor INTO  @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow
					END
					CLOSE precedingElementsCursor
					DEALLOCATE precedingElementsCursor
					
					IF (@fInvalidElements = 0) SET @iAction = 1
				END
		
				IF @iAction = 4 -- Or check
				BEGIN
					SET @fValidElements = 0		
					-- Check if any preceding steps have completed before submitting this step. 
		
					DECLARE precedingElementsCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT WE.ID,
						WE.type,
						WIS.status,
						WIS.decisionFlow
					FROM [dbo].[udfASRGetPrecedingWorkflowElements](@iElementID) PE
					INNER JOIN ASRSysWorkflowElements WE ON PE.ID = WE.ID
					INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
						AND WIS.instanceID = @iInstanceID

					OPEN precedingElementsCursor		
		
					FETCH NEXT FROM precedingElementsCursor INTO @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow
					WHILE (@@fetch_status = 0)
						AND (@fValidElements = 0)
					BEGIN
						IF (@iPrecedingElementType = 4) -- Decision
						BEGIN
							IF @iPrecedingElementStatus = 3 -- 3 = completed
							BEGIN
								SELECT @iCount = COUNT(*)
								FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iPrecedingElementFlow)
								WHERE ID = @iElementID
							
								IF @iCount > 0 SET @fValidElements = 1
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
										END

									SELECT @iCount = COUNT(*)
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
									WHERE ID = @iElementID
							
									IF @iCount > 0 SET @fValidElements = 1
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
											END

										SELECT @iCount = COUNT(*)
										FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iPrecedingElementID, @iTemp3)
										WHERE ID = @iElementID

										IF @iCount > 0 SET @fValidElements = 1
									END
								END
								ELSE
								BEGIN
									-- Preceding element must have status 3 (3 =Completed)
									IF @iPrecedingElementStatus = 3 SET @fValidElements = 1
								END
							END
						END

						FETCH NEXT FROM precedingElementsCursor INTO  @iPrecedingElementID, @iPrecedingElementType, @iPrecedingElementStatus, @iPrecedingElementFlow
					END
					CLOSE precedingElementsCursor
					DEALLOCATE precedingElementsCursor
		
					-- If all preceding steps have been completed submit the Or step.
					IF @fValidElements > '


	SET @sSPCode_2 = '0 
					BEGIN
						-- Cancel any preceding steps that are not completed as they are no longer required.
						EXEC [dbo].[spASRCancelPendingPrecedingWorkflowElements] @iInstanceID, @iElementID
		
						SET @iAction = 1
					END
				END
		
				IF @iAction = 1
				BEGIN
					EXEC [dbo].[spASRSubmitWorkflowStep] @iInstanceID, @iElementID, '''', '''', @sForms OUTPUT, @fSaveForLater OUTPUT
				END
		
				IF @iAction = 2
				BEGIN
					UPDATE ASRSysWorkflowInstanceSteps
					SET status = 2
					WHERE id = @iStepID
				END
		
				FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID
			END

			CLOSE stepsCursor
			DEALLOCATE stepsCursor

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
				END >= WE.timeoutFrequency

			OPEN timeoutCursor
			FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID
			WHILE (@@fetch_status = 0)
			BEGIN
				-- Set the step status to be Timeout
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 6, -- Timeout
					ASRSysWorkflowInstanceSteps.timeoutCount = isnull(ASRSysWorkflowInstanceSteps.timeoutCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID

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
						OR ASRSysWorkflowInstanceSteps.status = 8)
					
				-- Set activated Web Forms to be ''pending'' (to be done by the user)
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 2
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 2)
					
				-- Set activated Terminators to be ''completed''
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(), 
					ASRSysWorkflowInstanceSteps.completionCount = isnull(ASRSysWorkflowInstanceSteps.completionCount, 0) + 1
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSte'


	SET @sSPCode_3 = 'ps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 1)
					
				-- Count how many terminators have completed. ie. if the workflow has completed.
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type = 1
										
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @iInstanceID
					
					-- NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					-- is performed by a DELETE trigger on the ASRSysWorkflowInstances table.
				END

				FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID
			END

			CLOSE timeoutCursor
			DEALLOCATE timeoutCursor
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3)

	----------------------------------------------------------------------
	-- spASRGetWorkflowItemValues
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowItemValues]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowItemValues]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowItemValues]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRGetWorkflowItemValues]
			(
				@piElementItemID	integer,
				@piInstanceID	integer
			)
			AS
			BEGIN
				DECLARE 
					@iItemType			integer,
					@iResultType	integer,
					@sResult		varchar(8000),
					@fResult		bit,
					@dtResult		datetime,
					@fltResult		float,
					@iDefaultValueType		integer,
					@iCalcID				integer,
					@iLookupColumnID	integer,
					@sDefaultValue		varchar(8000),
					@sTableName			sysname,
					@sColumnName		sysname,
					@iDataType			integer,
					@sSelectSQL			varchar(8000),
					@iStatus			integer,
					@iElementID			integer,
					@sValue				varchar(8000),
					@sIdentifier		varchar(8000)
				DECLARE @dropdownValues TABLE([value] varchar(255))

				SELECT 			
					@iItemType = ASRSysWorkflowElementItems.itemType,
					@sDefaultValue = ASRSysWorkflowElementItems.inputDefault,
					@iLookupColumnID = ASRSysWorkflowElementItems.lookupColumnID,
					@iElementID = ASRSysWorkflowElementItems.elementID,
					@sIdentifier = ASRSysWorkflowElementItems.identifier,
					@iCalcID = isnull(ASRSysWorkflowElementItems.calcID, 0),
					@iDefaultValueType = isnull(ASRSysWorkflowElementItems.defaultValueType, 0)
				FROM ASRSysWorkflowElementItems
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID

				SELECT @iStatus = ASRSysWorkflowInstanceSteps.status
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @iElementID

				IF @iStatus = 7 -- Previously SavedForLater
				BEGIN
					SELECT @sValue = isnull(IVs.value, '''')
					FROM ASRSysWorkflowInstanceValues IVs
					WHERE IVs.instanceID = @piInstanceID
						AND IVs.elementID = @iElementID
						AND IVs.identifier = @sIdentifier

					SET @sDefaultValue = @sValue
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

						SET @sDefaultValue = 
							CASE
								WHEN @iResultType = 2 THEN convert(varchar(8000), @fltResult)
								WHEN @iResultType = 3 THEN 
									CASE 
										WHEN @fResult = 1 THEN ''TRUE''
										ELSE ''FALSE''
									END
								WHEN @iResultType = 4 THEN convert(varchar(100), @dtResult, 101)
								ELSE convert(varchar(8000), @sResult)
							END
					END
				END

				IF @iItemType = 15 -- OptionGroup
				BEGIN
					SELECT ASRSysWorkflowElementItemValues.value,
						CASE
							WHEN ASRSysWorkflowElementItemValues.value = @sDefaultValue THEN 1
							ELSE 0
						END
					FROM ASRSysWorkflowElementItemValues
					WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
					ORDER BY ASRSysWorkflowElementItemValues.sequence
				END

				IF @iItemType = 13 -- Dropdown
				BEGIN
					INSERT INTO @dropdownValues ([value])
						VALUES (null)

					INSERT INTO @dropdownValues ([value])
						SELECT ASRSysWorkflowElementItemValues.value
						FROM ASRSysWorkflowElementItemValues
						WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
						ORDER BY [sequence]

					SELECT [value],
						CASE
							WHEN [value] = @sDefaultValue THEN 1
							ELSE 0
						END
					FROM @dropdownValues 
				END
				
				IF (@iItemType = 14) -- Lookup
					AND (@iLookupColumnID > 0)
				BEGIN
					SELECT @sTableName = ASRSysTables.tableName,
						@sColumnName = ASRSysColumns.columnName,
						@iDataType = ASRSysColumns.dataType
					FROM ASRSysColumns
					INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
					WHERE ASRSysColumns.columnID = @iLookupColumnID

					IF @iDataType = 11 -- Date 
						AND UPPER(LTRIM(RTRIM(@sDefaultValue))) = ''NULL''
					BEGIN
						SET @sDefaultValue = ''''
					END

'


	SET @sSPCode_1 = '					SET @sSelectSQL = ''SELECT null AS [value], '' 
						+ CASE 
							WHEN @sDefaultValue = '''' THEN ''1''
							ELSE ''0''
						END
						+ '' UNION SELECT DISTINCT '' + @sTableName + ''.'' + @sColumnName + '' AS [value],''

					IF len(ltrim(rtrim(@sDefaultValue))) = 0 
					BEGIN
						SET @sSelectSQL = @sSelectSQL
							+ '' 0''
					END
					ELSE
					BEGIN
						SET @sSelectSQL = @sSelectSQL
							+ '' CASE''
							+ ''   WHEN '' + @sTableName + ''.'' + @sColumnName + '' = ''
							+ CASE
								WHEN (@iDataType = 12) -- Character
									OR (@iDataType = -1) -- WorkingPattern 
									OR (@iDataType = 11) -- Date 
									THEN '''''''' + REPLACE(@sDefaultValue, '''''''', '''''''''''') + ''''''''
								ELSE @sDefaultValue 
							END
							+ ''   THEN 1''
							+ ''   ELSE 0''
							+ '' END''
					END
					SET @sSelectSQL = @sSelectSQL
						+ '' FROM '' + @sTableName
						+ '' ORDER BY [value]''

					EXEC (@sSelectSQL)
				END
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

/* ------------------------------------------------------------- */
PRINT 'Step 7 of X - Add weekdays function'

	----------------------------------------------------------------------
	-- udfASRAddWeekdays
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfASRAddWeekdays]')
			AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfASRAddWeekdays]

	SET @sSPCode_0 = 'CREATE FUNCTION [dbo].[udfASRAddWeekdays] ()
		RETURNS integer
		AS
		BEGIN
			DECLARE @iDummy Int
			RETURN @iDummy
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER FUNCTION [dbo].[udfASRAddWeekdays]
		(
			@StartDate datetime, 
			@Duration int
		)
		RETURNS datetime
		AS
		BEGIN
		
			DECLARE @ReturnDate datetime
		
			IF NULLIF(@Duration, 0) IS NULL	
				RETURN @StartDate
		
			SELECT @ReturnDate = DATEADD(d,
								CASE DATEPART(dw,@StartDate) 
								WHEN 7 THEN 2 
								WHEN 1 THEN 1 
								ELSE 0 END,	@StartDate)
								+(DATEPART(dw,DATEADD(d,
									CASE DATEPART(dw,@StartDate) 
									WHEN 7 THEN 2 
									WHEN 1 THEN 1 
									ELSE 0 END,@StartDate))-2+@Duration)%5
								+((DATEPART(dw,DATEADD(d,
									CASE DATEPART(dw,@StartDate) 
									WHEN 7 THEN 2 
									WHEN 1 THEN 1 
									ELSE 0 END,@StartDate))-2+@Duration)/5)*7
								-(DATEPART(dw,DATEADD(d,
									CASE DATEPART(dw,@StartDate) 
									WHEN 7 THEN 2 
									WHEN 1 THEN 1 
									ELSE 0 END,@StartDate))-2)
		
			IF @ReturnDate IS NULL
				RETURN @StartDate
		
			RETURN @ReturnDate
		END'

	EXECUTE (@sSPCode_0)	

/* ------------------------------------------------------------- */
PRINT 'Step 8 of X - SQL 2008 Compatibility'

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

	----------------------------------------------------------------------
	-- spASRDefragIndexes
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRDefragIndexes]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRDefragIndexes]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRDefragIndexes]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRDefragIndexes]
			(@maxfrag DECIMAL)
		AS
			BEGIN
		
			SET NOCOUNT ON
		
			DECLARE @tablename VARCHAR (128)
			DECLARE @execstr VARCHAR (255)
			DECLARE @objectid INT
			DECLARE @objectowner VARCHAR(255)
			DECLARE @indexid INT
			DECLARE @frag DECIMAL
			DECLARE @indexname CHAR(255)
			DECLARE @dbname sysname
			DECLARE @tableid INT
			DECLARE @tableidchar VARCHAR(255)
			DECLARE @sSQLVersion int
		
			SELECT @sSQLVersion = dbo.udfASRSQLVersion()
		
			-- Checking fragmentation
			DECLARE tables CURSOR FOR
				SELECT convert(varchar,so.id)
				FROM sysobjects so
				JOIN sysindexes si ON so.id = si.id
				WHERE so.type =''U'' AND si.indid < 2 AND si.rows > 0
		
			-- Create the temporary table to hold fragmentation information
			CREATE TABLE #fraglist (
				ObjectName CHAR (255),
				ObjectId INT,
				IndexName CHAR (255),
				IndexId INT,
				Lvl INT,
				CountPages INT,
				CountRows INT,
				MinRecSize INT,
				MaxRecSize INT,
				AvgRecSize INT,
				ForRecCount INT,
				Extents INT,
				ExtentSwitches INT,
				AvgFreeBytes INT,
				AvgPageDensity INT,
				ScanDensity DECIMAL,
				BestCount INT,
				ActualCount INT,
				LogicalFrag DECIMAL,
				ExtentFrag DECIMAL)
		
			-- Open the cursor
			OPEN tables
		
			-- Loop through all the tables in the database running dbcc showcontig on each one
			FETCH NEXT FROM tables INTO @tableidchar
		
			WHILE @@FETCH_STATUS = 0
			BEGIN
				-- Do the showcontig of all indexes of the table
				INSERT INTO #fraglist 
				EXEC (''DBCC SHOWCONTIG ('' + @tableidchar + '') WITH FAST, TABLERESULTS, ALL_INDEXES, NO_INFOMSGS'')
				FETCH NEXT FROM tables INTO @tableidchar
			END
		
			-- Close and deallocate the cursor
			CLOSE tables
			DEALLOCATE tables
		
			-- Begin Stage 2: (defrag) declare cursor for list of indexes to be defragged
			DECLARE indexes CURSOR FOR
			SELECT ObjectName, ObjectOwner = user_name(so.uid), ObjectId, IndexName, ScanDensity
			FROM #fraglist f
			JOIN sysobjects so ON f.ObjectId=so.id
			WHERE ScanDensity <= @maxfrag
				AND INDEXPROPERTY (ObjectId, IndexName, ''IndexDepth'') > 0
		
			-- Open the cursor
			OPEN indexes
		
			-- Loop through the indexes
			FETCH NEXT
			FROM indexes
			INTO @tablename, @objectowner, @objectid, @indexname, @frag
		
			WHILE @@FETCH_STATUS = 0
			BEGIN
				SET QUOTED_IDENTIFIER ON
		
				IF (@sSQLVersion = 8)
					SET @execstr = ''DBCC DBREINDEX ('''''' + RTRIM(@objectowner) + ''.'' + RTRIM(@tablename)  + '''''' , '' + RTRIM(@indexname) +'')''
				ELSE
					SET @execstr = ''ALTER INDEX '' +  RTRIM(@indexname) + '' ON '' + RTRIM(@objectowner) + ''.'' + RTRIM(@tablename) + '' REBUILD''
		
				EXEC (@execstr)
		
				SET QUOTED_IDENTIFIER OFF
		
				FETCH NEXT FROM indexes INTO @tablename, @objectowner, @objectid, @indexname, @frag
			END
		
			-- Close and deallocate the cursor
			CLOSE indexes
			DEALLOCATE indexes
		
		
			DROP TABLE #FragList
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
		
			SET NOCOUNT ON
		
			DECLARE @sSQLVersion int
			DECLARE @Mode smallint
		
			IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
			BEGIN
				EXEC sp_ASRIntCheckPolls
			END
		
			SELECT @sSQLVersion = dbo.udfASRSQLVersion()
			SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = ''ProcessAccount'' AND [SettingKey] = ''Mode''
			IF @@ROWCOUNT = 0 SET @Mode = 0
			
			IF ((@Mode = 1 OR @Mode = 2) AND @sSQLVersion > 8) AND (NOT IS_SRVROLEMEMBER(''sysadmin'') = 1)
			BEGIN
				SELECT @piCount = dbo.[udfASRNetCountCurrentUsersInApp](APP_NAME())
			END
			ELSE
			BEGIN
		
				SELECT @piCount = COUNT(p.Program_Name)
				FROM     master..sysprocesses p
				JOIN     master..sysdatabases d
				  ON     d.dbid = p.dbid
				WHERE    p.program_name = APP_NAME()
				  AND    d.name = db_name()
				GROUP BY p.program_name
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
			@psLoginName	varchar(8000)
		)
		AS
		BEGIN
		
			DECLARE @sSQLVersion int
			DECLARE @Mode smallint
		
			IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
			BEGIN
				EXEC sp_ASRIntCheckPolls
			END
		
			SELECT @sSQLVersion = dbo.udfASRSQLVersion()
			SELECT @Mode = [SettingValue] FROM ASRSysSystemSettings WHERE [Section] = ''ProcessAccount'' AND [SettingKey] = ''Mode''
			IF @@ROWCOUNT = 0 SET @Mode = 0
			
			IF ((@Mode = 1 OR @Mode = 2) AND @sSQLVersion > 8) AND (NOT IS_SRVROLEMEMBER(''sysadmin'') = 1)
			BEGIN
				SELECT @iLoginCount = dbo.[udfASRNetCountCurrentLogins](@psLoginName)
			END
			ELSE
			BEGIN
		
				SELECT @iLoginCount = COUNT(*)
				FROM master..sysprocesses p
				WHERE p.program_name LIKE ''HR Pro%''
					AND p.program_name NOT LIKE ''HR Pro Workflow%''
		            AND p.program_name NOT LIKE ''HR Pro Outlook%''
		            AND p.program_name NOT LIKE ''HR Pro Server.Net%''
				    AND p.loginame = @psLoginName
			END
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- sp_ASRIntCheckLogin
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRIntCheckLogin]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRIntCheckLogin]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRIntCheckLogin]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[sp_ASRIntCheckLogin] (
			@piSuccessFlag			integer			OUTPUT,
			@psErrorMessage			varchar(8000)	OUTPUT,
			@piMinPassordLength		integer			OUTPUT,
			@psIntranetAppVersion	varchar(50),
			@piPasswordLength		integer,
			@piUserType				integer			OUTPUT
		)
		AS
		BEGIN
			/* Check that the current user is okay to login. */
			/* 	@pfLoginOK 	= 0 if the login was NOT okay
						= 1 if the login was okay (without warnings)
						= 2 if the login was okay (but the user''s password has expired)
				@psErrorMessage is the description of the login failure if @pfLoginOK = 0
				@piMinPassordLength is the configured minimum password length
				@psIntranetAppVersion is the intranet application version passed into the stored procedure (set as a session variable in the global.asa file. 
				@piPasswordLength is the length of the user''s current password. 
			*/
			SET NOCOUNT ON
			
			DECLARE @iSysAdminRoles				integer,
				@sLockUser						sysname,
				@sHostName						varchar(8000),
				@sLoginName						varchar(8000),
				@sProgramName					varchar(8000),
				@sRoleName						sysname,
				@source 						varchar(30),
				@desc 							varchar(200),
				@fIntranetEnabled				bit,
				@sIntranetDBVersion				varchar(50),
				@sIntranetDBMajor				varchar(50),
				@sIntranetDBMinor				varchar(50),
				@sIntranetDBRevision			varchar(50),
				@sIntranetAppMajor				varchar(50),
				@sIntranetAppMinor				varchar(50),
				@sIntranetAppRevision			varchar(50),
				@sMinIntranetVersion			varchar(50),
				@sMinIntranetMajor				varchar(50),
				@sMinIntranetMinor				varchar(50),
				@sMinIntranetRevision			varchar(50),
				@iPosition1 					integer,
				@iPosition2 					integer,
				@fValidIntranetAppVersion		bit,
				@fValidIntranetDBVersion		bit,
				@fValidMinIntranetVersion		bit,
				@iMinPasswordLength				integer,
				@iChangePasswordFrequency		integer,
				@sChangePasswordPeriod			varchar(1),
				@dtPasswordLastChanged			datetime,
				@fPasswordForceChange			bit,
				@sDomain						varchar(8000),
				@iCount							integer,
				@iPriority						integer,
				@sDescription					varchar(8000),
				@sLockTime						varchar(8000),
				@sValue							varchar(8000),
				@iValue							integer,
				@iFullUsers						integer,
				@iSSUsers						integer,
				@iSSIUsers						integer,
				@iTemp							integer,
				@fSelfService					bit, 
				@fValidSYSManagerVersion		bit,
				@sSYSManagerMajor				varchar(50),
				@sSYSManagerMinor				varchar(50),
				@sSYSManagerVersion				varchar(50),
				@sActualUserName				sysname, 
				@iActualUserGroupID				integer,
				@iFullIntItemID					integer,
				@iSSIntItemID					integer,
				@iSSIIntItemID					integer,
				@iSID							binary(85),
				@sSQLVersion					int,
				@sLockMessage					varchar(200),
				@fNewSettingFound				bit,
				@fOldSettingFound				bit
			SET @piSuccessFlag = 1
			SET @psErrorMessage = ''''
			SET @piMinPassordLength = 0
			SET @piUserType = 0
			SET @iFullUsers = 0
			SET @iSSUsers = 0
			SET @iSSIUsers = 0
			
			SET @fSelfService = 0
			IF APP_NAME() = ''HR Pro Self-service Intranet''
			BEGIN
				SET @fSelfService = 1
			END
			
			/*'' Check if the current user is a SQL Server System Administrator.
			We do not allow these users to login to the intranet module. */
			IF current_user = ''dbo''
			BEGIN
				SET @piSuccessFlag = 0
				SET @psErrorMessage = ''SQL Server system administrators cannot use the intranet module.''
			END
			ELSE
			BEGIN
				/* Fault 3901 */
				SELECT @iSysAdminRoles = sysAdmin + securityAdmin + serverAdmin + setupAdmin + processAdmin + diskAdmin + dbCreator
				FROM master..syslogins
				WHERE name = system_user
				IF @iSysAdminRoles > 0 
				BEGIN
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''Users assigned to fixed SQL Server roles cannot use the intranet module.''
				END
			END
			/* Check if anyone has locked the system. */
			IF @piSucc'


	SET @sSPCode_1 = 'essFlag = 1
			BEGIN
				CREATE TABLE #tmpSysProcess1 (hostname nvarchar(50), loginname nvarchar(50), program_name nvarchar(50), hostprocess int, sid binary(86), login_time datetime, spid smallint, uid smallint)
				INSERT #tmpSysProcess1 EXEC dbo.spASRGetCurrentUsers
			
				SELECT TOP 1 @iPriority = ASRSysLock.priority,
					@sLockUser = ASRSysLock.username,
					@sLockTime = convert(varchar(8000), ASRSysLock.lock_time, 100),
					@sHostName = ASRSysLock.hostname,
					@sDescription = ASRSysLock.description
				FROM ASRSysLock
				LEFT OUTER JOIN #tmpSysProcess1 syspro 
					ON ASRSysLock.spid = syspro.spid AND ASRSysLock.login_time = syspro.login_time
				WHERE priority = 2 
					OR syspro.spid IS not null
				ORDER BY priority
				IF (NOT @iPriority IS NULL) AND (@iPriority <> 3)
				BEGIN
					/* Get the lock message set in HR Pro System Manager */
					SET @sLockMessage = ''''
					EXEC sp_ASRIntGetSystemSetting ''messaging'', ''lockmessage'', ''lockmessage'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
					
					IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1) ) AND LTRIM(RTRIM(@sValue)) <> ''''
					BEGIN
						SET @sLockMessage = @sValue + ''<BR><BR>''
					END
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''The database has been locked.<P>'' + 
						CASE @iPriority
						WHEN 2 THEN
							@sLockMessage 
						ELSE ''''
						END
					    + ''User :  '' + @sLockUser + ''<BR>'' +
						  ''Date/Time :  '' + @sLockTime +  ''<BR>'' +
						  ''Machine :  '' + @sHostName +  ''<BR>'' +
						  ''Type :  '' + @sDescription
				END
				DROP TABLE #tmpSysProcess1
			END
			IF @piSuccessFlag = 1
			BEGIN
			
				/* Get the current HR Pro System Manager version */
				SET @sSYSManagerVersion = ''''
				exec sp_ASRIntGetSystemSetting ''database'', ''version'', ''version'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				
				IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
				BEGIN
					SET @sSYSManagerVersion = @sValue
				END
				/* Get the intranet version. */
				SET @sIntranetDBVersion = ''''
				IF @fSelfService = 0
				BEGIN
					exec dbo.sp_ASRIntGetSystemSetting ''intranet'', ''version'', ''intranetVersion'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				END
				ELSE
				BEGIN
					exec dbo.sp_ASRIntGetSystemSetting ''ssintranet'', ''version'', '''', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				END
				IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
				BEGIN
					SET @sIntranetDBVersion = @sValue
				END
				/* Get the minimum intranet version. */
				SET @sMinIntranetVersion = ''''
				IF @fSelfService = 0
				BEGIN
					exec dbo.sp_ASRIntGetSystemSetting ''intranet'', ''minimum version'', ''minIntranetVersion'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				END
				ELSE
				BEGIN
					exec dbo.sp_ASRIntGetSystemSetting ''ssintranet'', ''minimum version'', '''', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				END
				
				IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
				BEGIN
					SET @sMinIntranetVersion = @sValue
				END
				/* Get the minimum password length. */
				SET @iMinPasswordLength = 0
				exec dbo.sp_ASRIntGetSystemSetting ''password'', ''minimum length'', ''minimumPasswordLength'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
				BEGIN
					SET @iMinPasswordLength = convert(integer, @sValue)
				END
				SET @piMinPassordLength = @iMinPasswordLength
				/* Get the password change frequency. */
				SET @iChangePasswordFrequency = 0
				exec dbo.sp_ASRIntGetSystemSetting ''password'', ''change frequency'', ''changePasswordFrequency'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1'


	SET @sSPCode_2 = ') 
				BEGIN
					SET @iChangePasswordFrequency = convert(integer, @sValue)
				END
				/* Get the password change period. */
				SET @sChangePasswordPeriod = ''''
				exec dbo.sp_ASRIntGetSystemSetting ''password'', ''change period'', ''changePasswordFrequency'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				IF (@fNewSettingFound = 1) OR (@fOldSettingFound = 1) 
				BEGIN
					SET @sChangePasswordPeriod = UPPER(@sValue)
				END
			END
			/* Check the database version is the right one for the application version. */
			IF @piSuccessFlag = 1
			BEGIN
				/* Extract the Intranet application version parts from the given version string. */	
				SET @fValidIntranetAppVersion = 1
				SET @iPosition1 = charindex(''.'', @psIntranetAppVersion)
				IF @iPosition1 = 0 SET @fValidIntranetAppVersion = 0
				IF @fValidIntranetAppVersion = 1
				BEGIN
					SET @iPosition2 = charindex(''.'', @psIntranetAppVersion, @iPosition1 + 1)
					IF @iPosition2 = 0 SET @fValidIntranetAppVersion = 0
				END
				IF @fValidIntranetAppVersion = 1
				BEGIN
					SET @sIntranetAppMajor = left(@psIntranetAppVersion, @iPosition1 - 1)
					SET @sIntranetAppMinor = substring(@psIntranetAppVersion, @iPosition1 + 1, @iPosition2 - @iPosition1 - 1)
					SET @sIntranetAppRevision = substring(@psIntranetAppVersion, @iPosition2 + 1, len(@psIntranetAppVersion) - @iPosition2)
				END
				ELSE
				BEGIN
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''Invalid intranet application version.''
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				/* Extract the Intranet database version parts from the version string. */	
				SET @fValidIntranetDBVersion = 1
				SET @iPosition1 = charindex(''.'', @sIntranetDBVersion)
				IF @iPosition1 = 0 SET @fValidIntranetDBVersion = 0
				IF @fValidIntranetDBVersion = 1
				BEGIN
					SET @iPosition2 = charindex(''.'', @sIntranetDBVersion, @iPosition1 + 1)
					IF @iPosition2 = 0 SET @fValidIntranetDBVersion = 0
				END
				IF @fValidIntranetDBVersion = 1
				BEGIN
					SET @sIntranetDBMajor = left(@sIntranetDBVersion, @iPosition1 - 1)
					SET @sIntranetDBMinor = substring(@sIntranetDBVersion, @iPosition1 + 1, @iPosition2 - @iPosition1 - 1)
					SET @sIntranetDBRevision = substring(@sIntranetDBVersion, @iPosition2 + 1, len(@sIntranetDBVersion) - @iPosition2)
				END
				ELSE
				BEGIN
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''Invalid intranet database version.''
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				/* Extract the Minimum Intranet version parts from the version string. */	
				SET @fValidMinIntranetVersion = 1
				SET @iPosition1 = charindex(''.'', @sMinIntranetVersion)
				IF @iPosition1 = 0 SET @fValidMinIntranetVersion = 0
				IF @fValidMinIntranetVersion = 1
				BEGIN
					SET @iPosition2 = charindex(''.'', @sMinIntranetVersion, @iPosition1 + 1)
					IF @iPosition2 = 0 SET @fValidMinIntranetVersion = 0
				END
				IF @fValidMinIntranetVersion = 1
				BEGIN
					SET @sMinIntranetMajor = left(@sMinIntranetVersion, @iPosition1 - 1)
					SET @sMinIntranetMinor = substring(@sMinIntranetVersion, @iPosition1 + 1, @iPosition2 - @iPosition1 - 1)
					SET @sMinIntranetRevision = substring(@sMinIntranetVersion, @iPosition2 + 1, len(@sMinIntranetVersion) - @iPosition2)
				END
			END
			
			/* Check the System Manager database version is the right one for the intranet version. */
			IF @piSuccessFlag = 1
			BEGIN
				/* Extract the System Manager database version parts from the given version string. */	
				SET @fValidSYSManagerVersion = 1
				SET @iPosition1 = charindex(''.'', @sSYSManagerVersion)
				IF @iPosition1 = 0 SET @fValidSYSManagerVersion = 0
				IF @fValidSYSManagerVersion = 1
				BEGIN
					SET @sSYSManagerMajor = left(@sSYSManagerVersion, @iPosition1 - 1)
					SET @sSYSManagerMinor = substring(@sSYSManagerVersion, @iPosition1 + 1, len(@sSYSManagerVersion) - @iPosition1)
				END
				E'


	SET @sSPCode_3 = 'LSE
				BEGIN
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''Invalid System Manager database version.''
				END
			END
			
			IF @piSuccessFlag = 1
			BEGIN
				/* Check the application version against the one for the current database. */
				IF (convert(integer, @sIntranetAppMajor) < convert(integer, @sIntranetDBMajor)) 
					OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sIntranetDBMajor)) AND (convert(integer, @sIntranetAppMinor) < convert(integer, @sIntranetDBMinor))) 
					OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sIntranetDBMajor)) AND (convert(integer, @sIntranetAppMinor) = convert(integer, @sIntranetDBMinor)) AND (convert(integer, @sIntranetAppRevision) < convert(integer, @sIntranetDBRevision))) 
				BEGIN
					/* Application is too old for the database. */
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''The intranet application is out of date.'' 
														+ ''<BR>Please ask the System Administrator to update the intranet application.''
														+ ''<BR><BR>''
														+ ''Database Name : '' + db_name()
														+ ''<BR><BR>''
														+ ''HR Pro System Manager Version : '' + @sSYSManagerVersion
														+ ''<BR><BR>''
														+ ''HR Pro Intranet Database Version : '' + @sIntranetDBVersion
														+ ''<BR><BR>''
														+ ''HR Pro Intranet Application Version : '' + @sIntranetAppMajor + ''.'' + @sIntranetAppMinor + ''.'' + @sIntranetAppRevision				
														
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				/* Check the application version against the one for the current database. */
				IF (convert(integer, @sIntranetAppMajor) > convert(integer, @sIntranetDBMajor)) 
					OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sIntranetDBMajor)) AND (convert(integer, @sIntranetAppMinor) > convert(integer, @sIntranetDBMinor))) 
					OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sIntranetDBMajor)) AND (convert(integer, @sIntranetAppMinor) = convert(integer, @sIntranetDBMinor)) AND (convert(integer, @sIntranetAppRevision) > convert(integer, @sIntranetDBRevision))) 
				BEGIN
					/* Database is too old for the appplication. */
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''The database is out of date.'' 
														+ ''<BR>Please ask the System Administrator to update the database for use with version '' + @sIntranetAppMajor + ''.'' + @sIntranetAppMinor + ''.'' + @sIntranetAppRevision + '' of the intranet.''
														+ ''<BR><BR>''
														+ ''Database Name : '' +  db_name()
														+ ''<BR><BR>''
														+ ''HR Pro System Manager Version : '' + @sSYSManagerVersion
														+ ''<BR><BR>''
														+ ''HR Pro Intranet Database Version : '' + @sIntranetDBVersion
														+ ''<BR><BR>''
														+ ''HR Pro Intranet Application Version : '' + @sIntranetAppMajor + ''.'' + @sIntranetAppMinor + ''.'' + @sIntranetAppRevision				
														
					IF (convert(integer, @sIntranetAppMajor) > convert(integer, @sSYSManagerMajor)) 
							OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sSYSManagerMajor)) AND (convert(integer, @sIntranetAppMinor) > convert(integer, @sSYSManagerMinor))) 
					BEGIN
						SET @psErrorMessage = @psErrorMessage + ''<BR><BR>''
																									+ ''<FONT COLOR="Red"><B>Please note that the System Manager version also requires updating to version '' + @sIntranetAppMajor + ''.'' + @sIntranetAppMinor + ''.</B></FONT>'' 
					END					
														
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				/* Check the application version against the one for the current database. */
				IF (convert(integer, @sIntranetAppMajor) > convert(integer, @sSYSManagerMajor)) 
					OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sSYSManagerMajor)) AND (convert(integer, @sIntranetAppMinor) > convert(integer, @sSYSManagerMinor))) 
				BEGIN
					'


	SET @sSPCode_4 = '/* Database is too old for the appplication. */
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''The database is out of date.'' 
						+ ''<BR>Please ask the System Administrator to update the System Manager version to '' + @sIntranetAppMajor + ''.'' + @sIntranetAppMinor + ''.''
						+ ''<BR><BR>''
						+ ''Database Name : '' + db_name()
						+ ''<BR><BR>''
						+ ''HR Pro System Manager Version : '' + @sSYSManagerVersion
						+ ''<BR><BR>''
						+ ''HR Pro Intranet Database Version : '' + @sIntranetDBVersion
						+ ''<BR><BR>''
						+ ''HR Pro Intranet Application Version : '' + @sIntranetAppMajor + ''.'' + @sIntranetAppMinor + ''.'' + @sIntranetAppRevision				
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				EXEC sp_ASRIntGetSystemSetting ''platform'', ''SQLServerVersion'', ''SQLServerVersion'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				
				IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1) ) AND SUBSTRING(LTRIM(RTRIM(@sValue)),1,1) <> @sSQLVersion
				BEGIN
					/* Microsoft SQL Version has been upgraded */
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''The Microsoft SQL Version has been upgraded.'' 
						+ ''<BR>Please ask the System Administrator to save the update in the System Manager.''
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				EXEC sp_ASRIntGetSystemSetting ''platform'', ''DatabaseName'', ''DatabaseName'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1) ) AND UPPER(LTRIM(RTRIM(@sValue))) <> UPPER(DB_NAME())
				BEGIN
					/* The database name changed */
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''The database name has changed.'' 
						+ ''<BR>Please ask the System Administrator to save the update in the System Manager.''
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				EXEC sp_ASRIntGetSystemSetting ''platform'', ''ServerName'', ''ServerName'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
				
				IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1))
				BEGIN
					IF LTRIM(RTRIM(@sValue)) = ''.'' SELECT @sValue = @@SERVERNAME
					IF UPPER(@sValue) <> UPPER(@@SERVERNAME)
					BEGIN
						/* The database has moved to a different Microsoft SQL Server */
						SET @piSuccessFlag = 0
						SET @psErrorMessage = ''The database has moved to a different Microsoft SQL Server.'' 
							+ ''<BR>Please ask the System Administrator to save the update in the System Manager.''
					END
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				EXEC sp_ASRIntGetSystemSetting ''database'', ''refreshstoredprocedures'', ''refreshstoredprocedures'', @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
					
				IF ((@fNewSettingFound = 1) OR (@fOldSettingFound = 1) ) AND LTRIM(RTRIM(@sValue)) = 1
				BEGIN
					/* Database is too old for the appplication. */
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''The database is out of date.'' 
						+ ''<BR>Please ask the System Administrator to save the update in the System Manager.''
						+ ''<BR><BR>''
						+ ''Database Name : '' + db_name()
						+ ''<BR><BR>''
						+ ''HR Pro System Manager Version : '' + @sSYSManagerVersion
						+ ''<BR><BR>''
						+ ''HR Pro Intranet Database Version : '' + @sIntranetDBVersion
						+ ''<BR><BR>''
						+ ''HR Pro Intranet Application Version : '' + @sIntranetAppMajor + ''.'' + @sIntranetAppMinor + ''.'' + @sIntranetAppRevision	
				END
			END
			IF (@piSuccessFlag = 1) AND (@fValidMinIntranetVersion = 1)
			BEGIN
				/* Check the application version against the minimum one for the current database. */
				IF (convert(integer, @sIntranetAppMajor) < convert(integer, @sMinIntranetMajor)) 
					OR ((convert(integer, @sIntranetAppMajor) = convert(integer, @sMinIntranetMajor)) AND (convert(integer, @sIntranetAppMinor) < convert(integer, @sMinIntranetMinor))) 
					O'


	SET @sSPCode_5 = 'R ((convert(integer, @sIntranetAppMajor) = convert(integer, @sMinIntranetMajor)) AND (convert(integer, @sIntranetAppMinor) = convert(integer, @sMinIntranetMinor)) AND (convert(integer, @sIntranetAppRevision) < convert(integer, @sMinIntranetRevision))) 
				BEGIN
					/* Application is older than the minimum required */
					SET @piSuccessFlag = 0
					--SET @psErrorMessage = ''The intranet application is out of date. You require version '' + @sMinIntranetVersion + '' or later. Contact your administrator to update it.''
					SET @psErrorMessage = ''The intranet application is out of date.'' 
														+ ''<BR>Please ask the System Administrator to update the intranet application.''
														+ ''<BR><BR>''
														+ ''Database Name : '' + db_name()
														+ ''<BR><BR>''
														+ ''HR Pro System Manager Version : '' + @sSYSManagerVersion
														+ ''<BR><BR>''
														+ ''HR Pro Intranet Database Version : '' + @sIntranetDBVersion
														+ ''<BR><BR>''
														+ ''HR Pro Intranet Application Version : '' + @sIntranetAppMajor + ''.'' + @sIntranetAppMinor + ''.'' + @sIntranetAppRevision				
				END
			END
			-- Get licence details
			IF @piSuccessFlag = 1
			BEGIN
				EXEC dbo.spASRIntGetLicenceInfo @fSelfService, @piSuccessFlag OUTPUT,
							@fIntranetEnabled OUTPUT, @iSSUsers OUTPUT,
							@iFullUsers OUTPUT,	@iSSIUsers OUTPUT,
							@psErrorMessage OUTPUT
			END
			-- Check that the user belongs to a valid role in the selected database.
			IF @piSuccessFlag = 1
			BEGIN
				EXEC dbo.spASRIntGetActualUserDetails
					@sActualUserName OUTPUT,
					@sRoleName OUTPUT,
					@iActualUserGroupID OUTPUT
							        
				IF @sRoleName IS NULL
				BEGIN
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''The  user is not a member of any HR Pro user group.''
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				/* Check that the user is permitted to use the Intranet module. */
				/* First check that this permission exists in the current version of HR Pro. */
				SELECT @iSSIIntItemID = ASRSysPermissionItems.itemID
				FROM ASRSysPermissionItems
				INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
				WHERE ASRSysPermissionItems.itemKey = ''SSINTRANET''
					AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS''
				IF @iSSIIntItemID IS NULL SET @iSSIIntItemID = 0
						
				IF @fSelfService = 1
				BEGIN
					/* The permission does exist in the current version so check if the user is granted this permission. */
					SELECT @iCount = count(ItemID)
					FROM ASRSysGroupPermissions 
					WHERE ASRSysGroupPermissions.itemID = @iSSIIntItemID
						AND ASRSysGroupPermissions.groupName = @sRoleName
						AND ASRSysGroupPermissions.permitted = 1
					IF @iCount = 0
					BEGIN				
						SET @piSuccessFlag = 0
						SET @psErrorMessage = ''You are not permitted to use the Self-service Intranet module with this user name.''
					END
				END
				ELSE
				BEGIN
					SELECT @iFullIntItemID = ASRSysPermissionItems.itemID
					FROM ASRSysPermissionItems
					INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					WHERE ASRSysPermissionItems.itemKey = ''INTRANET''
						AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS''
					IF @iFullIntItemID IS NULL SET @iFullIntItemID = 0
				
					SELECT @iSSIntItemID = ASRSysPermissionItems.itemID
					FROM ASRSysPermissionItems
					INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
					WHERE ASRSysPermissionItems.itemKey = ''INTRANET_SELFSERVICE''
						AND ASRSysPermissionCategories.categoryKey = ''MODULEACCESS''
					IF @iSSIntItemID IS NULL SET @iSSIntItemID = 0
				
					IF @iFullIntItemID > 0
					BEGIN
						/* The permission does exist in the current version '


	SET @sSPCode_6 = 'so check if the user is granted this permission. */
						SELECT @iCount = count(ItemID)
						FROM ASRSysGroupPermissions 
						WHERE ASRSysGroupPermissions.itemID = @iFullIntItemID
							AND ASRSysGroupPermissions.groupName = @sRoleName
							AND ASRSysGroupPermissions.permitted = 1
						
						IF @iCount = 0
						BEGIN
							IF @iSSIntItemID > 0
							BEGIN
								/* The permission does exist in the current version so check if the user is granted this permission. */
								SELECT @iCount = count(ItemID)
								FROM ASRSysGroupPermissions 
								WHERE ASRSysGroupPermissions.itemID = @iSSIntItemID
									AND ASRSysGroupPermissions.groupName = @sRoleName
									AND ASRSysGroupPermissions.permitted = 1
								IF @iCount = 0
								BEGIN				
									SET @piSuccessFlag = 0
									SET @psErrorMessage = ''You are not permitted to use the Data Manager Intranet module with this user name.''
								END
								ELSE
								BEGIN
									SET @piUserType = 1
								END
							END
						END
					END
				END
			END
			IF @piSuccessFlag = 1
			BEGIN
				IF @fSelfService = 1
				BEGIN
					SET @iTemp = @iSSIUsers
				END
				ELSE
				BEGIN
					IF @piUserType = 1
					BEGIN
						SET @iTemp = @iSSUsers
					END
					ELSE
					BEGIN
						SET @iTemp = @iFullUsers
					END
				END
				SET @iValue = 0
				/* Don''t use uid as it sometimes is 0 when youdon''t expect it to be. */
				CREATE TABLE #tmpSysProcess2 (hostname nvarchar(50), loginname nvarchar(50), program_name nvarchar(50), hostprocess int, sid binary(86), login_time datetime, spid smallint, uid smallint)
				INSERT #tmpSysProcess2 EXEC dbo.spASRGetCurrentUsers
				DECLARE users_cursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT sid
					FROM #tmpSysProcess2
					WHERE program_name = APP_NAME()
				OPEN users_cursor
				FETCH NEXT FROM users_cursor INTO @iSID
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @fSelfService = 1
					BEGIN
						SET @iValue = @iValue + 1
					END
					ELSE
					BEGIN
						/* Check if the process is run by the same type of user as the current user. */
						/* Get the user''s group name. */
						SELECT @sRoleName = usg.name
						FROM sysusers usu
						left outer join
							(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
						WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) 
							AND (usg.issqlrole = 1 OR usg.uid is null) 
							AND usu.sid = @iSID 
							AND not (usg.name like ''ASRSys%'') 
							AND not (usg.name like ''db_owner'')
						IF @piUserType = 1
						BEGIN
							/* Self-service users. */
							IF @iSSIntItemID > 0
							BEGIN
								/* The permission does exist in the current version so check if the user is granted this permission. */
								SELECT @iCount = count(ItemID)
								FROM ASRSysGroupPermissions 
								WHERE ASRSysGroupPermissions.itemID = @iSSIntItemID
									AND ASRSysGroupPermissions.groupName = @sRoleName
									AND ASRSysGroupPermissions.permitted = 1
								IF @iCount > 0 SET @iValue = @iValue + 1
							END
						END
						ELSE
						BEGIN
							/* Full access users. */
							IF @iFullIntItemID > 0
							BEGIN
								/* The permission does exist in the current version so check if the user is granted this permission. */
								SELECT @iCount = count(*)
								FROM ASRSysGroupPermissions 
								WHERE ASRSysGroupPermissions.itemID = @iFullIntItemID
								AND ASRSysGroupPermissions.groupName = @sRoleName
								AND ASRSysGroupPermissions.permitted = 1
								IF @iCount > 0 SET @iValue = @iValue + 1
							END
						END
					END
					FETCH NEXT FROM users_cursor INTO @iSID
				END
				DROP TABLE #tmpSysProcess2
				
				CLOSE users_cursor
				DEALLOCATE users_cursor
				IF @iValue > @iTemp
				BEGIN
					SET @piSuccessFlag = 0
					SET @psErrorMessage = ''Unable to logon. You have reached the '


	SET @sSPCode_7 = 'maximum number of licensed '' + 
						CASE
							WHEN @fSelfService = 1 THEN ''Self-service Intranet''
							WHEN @piUserType = 1 THEN ''Data Manager Intranet (single record)''
							ELSE ''Data Manager Intranet (multiple record)''
						END +
						'' users.''
				END
			END
			/* Check if the password has expired */
			SELECT @sSQLVersion = dbo.udfASRSQLVersion()
			IF @piSuccessFlag = 1 AND @sSQLVersion < 9
			BEGIN
				SELECT @dtPasswordLastChanged = lastChanged, 
					@fPasswordForceChange = forceChange
				FROM ASRSysPasswords
				WHERE userName = system_user
				IF @dtPasswordLastChanged IS NULL
				BEGIN
					/* User not in the password table. So add them. */
					SET @dtPasswordLastChanged = GETDATE()
					SET @fPasswordForceChange = 0
					INSERT INTO ASRSysPasswords (username, lastChanged, forceChange)
					VALUES (LOWER(system_user), @dtPasswordLastChanged, @fPasswordForceChange)
				END
				ELSE
				BEGIN
					IF (@iMinPasswordLength <> 0) OR (@iChangePasswordFrequency <> 0) 
					BEGIN
						/* Check for minimum length. */
						IF (@iMinPasswordLength > @piPasswordLength) SET @fPasswordForceChange = 1
		    
						/* Check for Date last changed. */
						IF (@iChangePasswordFrequency > 0) AND (@fPasswordForceChange = 0)
						BEGIN
							IF @sChangePasswordPeriod = ''D'' 
							BEGIN
								IF DATEADD(day, @iChangePasswordFrequency, @dtPasswordLastChanged) <= GETDATE() SET @fPasswordForceChange = 1						
							END
							IF @sChangePasswordPeriod = ''W'' 
							BEGIN
								IF DATEADD(week, @iChangePasswordFrequency, @dtPasswordLastChanged) <= GETDATE() SET @fPasswordForceChange = 1						
							END
							IF @sChangePasswordPeriod = ''M'' 
							BEGIN
								IF DATEADD(month, @iChangePasswordFrequency, @dtPasswordLastChanged) <= GETDATE() SET @fPasswordForceChange = 1						
							END
							IF @sChangePasswordPeriod = ''Y'' 
							BEGIN
								IF DATEADD(year, @iChangePasswordFrequency, @dtPasswordLastChanged) <= GETDATE() SET @fPasswordForceChange = 1						
							END
						END
					END
				END
				IF @fPasswordForceChange = 1 SET @piSuccessFlag = 2
			END
			
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4
		+ @sSPCode_5
		+ @sSPCode_6
		+ @sSPCode_7)

	----------------------------------------------------------------------
	-- spASRGetDomainPolicy
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetDomainPolicy]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetDomainPolicy]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetDomainPolicy]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetDomainPolicy]
			(@LockoutDuration int OUTPUT,
			 @lockoutThreshold int OUTPUT,
			 @lockoutObservationWindow int OUTPUT,
			 @maxPwdAge int OUTPUT, 
			 @minPwdAge int OUTPUT,
			 @minPwdLength int OUTPUT, 
			 @pwdHistoryLength int OUTPUT, 
			 @pwdProperties int OUTPUT)
		AS
		BEGIN
		
			SET NOCOUNT ON
			
			DECLARE @objectToken int
			DECLARE @hResult int
			DECLARE @hResult2 int
			DECLARE @pserrormessage varchar(255)
			DECLARE @sSQLVersion int
		
			-- Initialise the variables
			SET @LockoutDuration = 0
			SET @lockoutThreshold  = 0
			SET @lockoutObservationWindow  = 0
			SET @maxPwdAge  = 0
			SET @minPwdAge  = 0
			SET @minPwdLength  = 0
			SET @pwdHistoryLength  = 0 
			SET @pwdProperties  = 0
		
			SELECT @sSQLVersion = dbo.udfASRSQLVersion()
		
			-- SQL2000 uses Server DLL, SQL2005+ uses server assembly
			IF @sSQLVersion = 8
			BEGIN
		
				/* Create Server DLL object */
				EXEC @hResult = sp_OACreate ''vbpHRProServer.clsDomainInfo'', @objectToken OUTPUT
				IF @hResult <> 0
				BEGIN
				  EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
				  SET @pserrormessage = ''HR Pro Server.dll not found''
				  RAISERROR (@pserrormessage,1,1)
				  EXEC sp_OADestroy @objectToken
				  RETURN 1
				END
		
				-- Populate the variables
				EXEC @hResult = sp_OAMethod @objectToken, ''getDomainPolicySettings'',@hResult2 OUTPUT, @LockoutDuration OUTPUT
						, @lockoutThreshold OUTPUT, @lockoutObservationWindow OUTPUT, @maxPwdAge OUTPUT
						, @minPwdAge OUTPUT, @minPwdLength OUTPUT, @pwdHistoryLength OUTPUT
						, @pwdProperties OUTPUT
		
				IF @hResult <> 0 
				BEGIN
				  EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
				  SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
				  RAISERROR (@pserrormessage,2,1)
				  EXEC sp_OADestroy @objectToken
				  RETURN 2
				END
		
				EXEC sp_OADestroy @objectToken
			END
			ELSE
			BEGIN
				EXEC sp_executesql N''EXEC spASRGetDomainPolicyFromAssembly
						@lockoutDuration OUTPUT, @lockoutThreshold OUTPUT,
						@lockoutObservationWindow OUTPUT, @maxPwdAge OUTPUT,
						@minPwdAge OUTPUT, @minPwdLength OUTPUT,
						@pwdHistoryLength OUTPUT, @pwdProperties OUTPUT''
					, N''@lockoutDuration int OUT, @lockoutThreshold int OUT,
						@lockoutObservationWindow int OUT, @maxPwdAge int OUT,
						@minPwdAge int OUT,	@minPwdLength int OUT,
						@pwdHistoryLength int OUT, @pwdProperties int OUT''
					, @LockoutDuration OUT, @lockoutThreshold OUT
					, @lockoutObservationWindow OUT, @maxPwdAge OUT
					, @minPwdAge OUT, @minPwdLength OUT
					, @pwdHistoryLength OUT, @pwdProperties OUT
			END
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetDomains
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetDomains]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetDomains]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetDomains]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetDomains]
			(@DomainString varchar(8000) OUTPUT)
		AS
		BEGIN
		
			SET NOCOUNT ON
		
			DECLARE @objectToken int
			DECLARE @hResult int
			DECLARE @hResult2 varchar(255)
			DECLARE @pserrormessage varchar(255)
			DECLARE @sSQLVersion int
		
			SELECT @sSQLVersion = dbo.udfASRSQLVersion()
		
			-- SQL2000 uses Server DLL, SQL2005+ uses server assembly
			IF @sSQLVersion = 8
			BEGIN
		
				-- Create Server DLL object
				EXEC @hResult = sp_OACreate ''vbpHRProServer.clsDomainInfo'', @objectToken OUTPUT
				IF @hResult <> 0
				BEGIN
				  EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
				  SET @pserrormessage = ''HR Pro Server.dll not found''
				  RAISERROR (@pserrormessage,1,1)
				  EXEC sp_OADestroy @objectToken
				  RETURN 1
				END
			
				-- Populate the variables
				EXEC @hResult = sp_OAMethod @objectToken, ''getDomains'', @hResult2 OUTPUT, @DomainString OUTPUT
			
				IF @hResult <> 0 
				BEGIN
				  EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
				  SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
				  RAISERROR (@pserrormessage,2,1)
				  EXEC sp_OADestroy @objectToken
				  RETURN 2
				END
			
				EXEC sp_OADestroy @objectToken
			END
			ELSE
			BEGIN
				SELECT @DomainString = dbo.udfASRNetGetDomains()
			END
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetWindowsUsers
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWindowsUsers]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWindowsUsers]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWindowsUsers]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetWindowsUsers]
		(
			@DomainName varchar(200),
			@UserString varchar(8000) OUTPUT
		)
		AS
		BEGIN
		
			DECLARE @objectToken int
			DECLARE @hResult int
			DECLARE @hResult2 varchar(255)
			DECLARE @pserrormessage varchar(255)
			DECLARE @sSQLVersion int
		
			SELECT @sSQLVersion = dbo.udfASRSQLVersion()
		
			-- SQL2000 uses Server DLL, SQL2005+ uses server assembly
			IF @sSQLVersion = 8
			BEGIN
			
				-- Create Server DLL object
				EXEC @hResult = sp_OACreate ''vbpHRProServer.clsDomainInfo'', @objectToken OUTPUT
				IF @hResult <> 0
				BEGIN
					EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
					SET @pserrormessage = ''HR Pro Server.dll not found''
					RAISERROR (@pserrormessage,1,1)
					EXEC sp_OADestroy @objectToken
					RETURN 1
				END
		
				-- Populate the variables
				EXEC @hResult = sp_OAMethod @objectToken, ''GetUsers'', @UserString OUTPUT, @DomainName
				IF @hResult <> 0 
				BEGIN
					EXEC sp_OAGetErrorInfo @objectToken, '''', @pserrormessage OUTPUT
					SET @pserrormessage = ''HR Pro Server.dll error (''+rtrim(ltrim(@pserrormessage))+'')''
					RAISERROR (@pserrormessage,2,1)
					EXEC sp_OADestroy @objectToken
				RETURN 2
				END
		
				EXEC sp_OADestroy @objectToken
			END
			ELSE
			BEGIN
				SELECT @UserString = dbo.udfASRNetGetUsers(@DomainName)
			END
		
			
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetWorkflowQueryString
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowQueryString]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowQueryString]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowQueryString]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowQueryString]
		(
			@piInstanceID	integer,
			@piElementID	integer,
			@psQueryString	varchar(8000)	output
		)
		AS
		BEGIN
			DECLARE
				@hResult		integer,
				@objectToken	integer,
				@sURL			varchar(8000),
				@sParam1		varchar(8000),
				@sDBName		sysname,
				@sSQLVersion	int
		
			SET @psQueryString = ''''
			SET @sSQLVersion = dbo.udfASRSQLVersion()
		
			SELECT @sURL = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_URL''
				
			IF upper(right(@sURL, 5)) <> ''.ASPX''
				AND right(@sURL, 1) <> ''/''
				AND len(@sURL) > 0
			BEGIN
				SET @sURL = @sURL + ''/''
			END
		
			SELECT @sParam1 = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_Web1''
		
			IF (len(@sURL) > 0)
			BEGIN
				SET @sDBName = db_name()
		
				IF @sSQLVersion = 8
				BEGIN			
					EXEC @hResult = sp_OACreate ''vbpHRProServer.clsWorkflow'', @objectToken OUTPUT
			
					IF (@hResult = 0) 
					BEGIN
						EXEC @hResult = sp_OAMethod @objectToken, ''GetQueryString'', @psQueryString OUTPUT, @piInstanceID, @piElementID, @sParam1, @@servername, @sDBName
						IF @hResult <> 0
						BEGIN
							SET @psQueryString = ''''
						END
			
						EXEC sp_OADestroy @objectToken
					END
				END
				ELSE
				BEGIN
					SELECT @psQueryString = dbo.[udfASRNetGetWorkflowQueryString]( @piInstanceID, @piElementID, @sParam1, @@servername, @sDBName)
				END
			
				IF len(@psQueryString) > 0
				BEGIN
					SET @psQueryString = @sURL + ''?'' + @psQueryString
				END
			END
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRUpdateStatistics
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRUpdateStatistics]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRUpdateStatistics]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRUpdateStatistics]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRUpdateStatistics]
		AS
		BEGIN
		
			SET NOCOUNT ON
		
			DECLARE @sTableName nvarchar(4000)
			DECLARE @sVarCommand nvarchar(4000)
		
			-- Checking fragmentation
			DECLARE tables CURSOR FOR
				SELECT so.[Name]
				FROM sysobjects so
				JOIN sysindexes si ON so.id = si.id
				WHERE so.type =''U'' AND si.indid < 2 AND si.rows > 0
				ORDER BY so.[Name]
		
			-- Open the cursor
			OPEN tables
		
			-- Loop through all the tables in the database running dbcc showcontig on each one
			FETCH NEXT FROM tables INTO @sTableName
		
			WHILE @@FETCH_STATUS = 0
			BEGIN
				SET @sVarCommand = ''UPDATE STATISTICS '' + @sTableName + '' WITH FULLSCAN''
				EXEC (@sVarCommand)
				FETCH NEXT FROM tables INTO @sTableName
			END
		
			-- Close and deallocate the cursor
			CLOSE tables
			DEALLOCATE tables
		
		END'

	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 9 of X - Save Settings Proc'

	----------------------------------------------------------------------
	-- spASRSaveSetting
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSaveSetting]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSaveSetting]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRSaveSetting]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRSaveSetting] (
			@psSection		varchar(8000),
			@psKey		varchar(8000),
			@psValue		varchar(8000)	
		)
		AS
		BEGIN
		
			/* Save the given system setting. */
			DELETE FROM ASRSysSystemSettings
			WHERE section = @psSection
				AND settingKey = @psKey
			
			INSERT INTO ASRSysSystemSettings 
				(section, settingKey, settingValue)
			VALUES (@psSection, @psKey, @psValue)
			
		END'

	EXECUTE (@sSPCode_0)

/* ------------------------------------------------------------- */

/* ------------------------------------------------------------- */
PRINT 'Step 10 of X - Modifying ASRSysColumns Table'

	--Does CalculateIfEmpty already exist?
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysColumns')
	and name = 'CalculateIfEmpty'

	if @iRecCount = 0
	BEGIN
	--Column does not exist, so create it
		SELECT @NVarCommand = 'ALTER TABLE [ASRSysColumns] ADD [CalculateIfEmpty] [bit] NOT NULL DEFAULT 0'
		EXEC sp_executesql @NVarCommand
	END
	ELSE
	--Column Exists, now check datatype
	BEGIN
		DECLARE @iColumnDataType	integer
		DECLARE @iBitDataType	integer

		SELECT @iColumnDataType = xtype FROM syscolumns
		where id = (select id from sysobjects where name = 'ASRSysColumns')
		and name = 'CalculateIfEmpty'
	
		SELECT @iBitDataType = xtype from systypes where name = 'BIT'

		if @iColumnDataType <> @iBitDataType	--type NOT bit, which is incorrect, so convert it to a bit type
		BEGIN
			--Replace all nulls with 0, and anything greater than 1 with 1
			SELECT @NVarCommand = 'UPDATE [ASRSysColumns] SET [CalculateIfEmpty] = 0 WHERE [CalculateIfEmpty] IS NULL'
			EXEC sp_executesql @NVarCommand
			
			SELECT @NVarCommand = 'UPDATE [ASRSysColumns] SET [CalculateIfEmpty] = 1 WHERE [CalculateIfEmpty] <> 0'
			EXEC sp_executesql @NVarCommand
			
			--Drop any constraints			
			declare @sConstraintName varchar(5000)

			select @sConstraintName = name from sysobjects where parent_obj = 
					(SELECT O.ID FROM sysobjects O, syscolumns C 
					WHERE C.ID=O.ID AND O.NAME='ASRSYSCOLUMNS' and O.XTYPE='U' and C.Name='CalculateIfEmpty')
					and ID =
					(SELECT c.cdefault FROM sysobjects O, syscolumns C 
					WHERE C.ID=O.ID AND O.NAME='ASRSYSCOLUMNS' and O.XTYPE='U' and C.Name='CalculateIfEmpty')

			
			SELECT @NVarCommand = 'ALTER TABLE [ASRSysColumns] DROP CONSTRAINT [' + @sConstraintName + ']'
			EXEC sp_executesql @NVarCommand
			--Alter column to a BIT type
			SELECT @NVarCommand = 'ALTER TABLE [ASRSysColumns] ALTER COLUMN [CalculateIfEmpty] BIT NOT NULL'
			EXEC sp_executesql @NVarCommand
		END
		ELSE
		BEGIN
			--Existing column is a BIT type, so make sure there are no NULLs
			SELECT @NVarCommand = 'UPDATE [ASRSysColumns] SET [CalculateIfEmpty] = 0 WHERE [CalculateIfEmpty] IS NULL'
			EXEC sp_executesql @NVarCommand

			
			--Change column to disallow NULLs in the future
			SELECT @NVarCommand = 'ALTER TABLE [ASRSysColumns] ALTER COLUMN [CalculateIfEmpty] BIT NOT NULL'
			EXEC sp_executesql @NVarCommand
		END
	END
/* ------------------------------------------------------------- */

/* ------------------------------------------------------------- */
PRINT 'Step 11 of X - Windows Authentication procedures'

	IF OBJECT_ID('ASRSysUserGroups', N'U') IS NULL	
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysUserGroups](
				[UserName] [varchar](1000) NOT NULL,
				[UserGroup] [varchar](1000) NOT NULL
			) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	----------------------------------------------------------------------
	-- spASRGetCurrentUsersInWindowsGroups
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersInWindowsGroups]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersInWindowsGroups]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersInWindowsGroups]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetCurrentUsersInWindowsGroups]
		(
			@psGroupNames VARCHAR(4000)
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
		
			CREATE TABLE #tblUsersInGroup
				(
				loginame varchar(256)
				,groupname varchar(256)
				)		
		
			DECLARE @tblGroups TABLE
				(
				groupname varchar(256)
				)
		
			DECLARE @iUserInGroup int,
					@loginame varchar(256),
					@IN varchar(4000), 
					@INGroup varchar(4000),
					@Pos int
		
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
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRGetActualUserDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetActualUserDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetActualUserDetails]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetActualUserDetails]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetActualUserDetails]
		(
				@psUserName sysname OUTPUT,
				@psUserGroup sysname OUTPUT,
				@piUserGroupID integer OUTPUT
		)
		AS
		BEGIN
			DECLARE @iFound		int
			DECLARE @sSQLVersion int
		
			SET @sSQLVersion = convert(int,convert(float,substring(@@version,charindex(''-'',@@version)+2,2)))
		
			SELECT @iFound = COUNT(*) 
			FROM sysusers usu 
			LEFT OUTER JOIN	(sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
			LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
			WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
				AND (usg.issqlrole = 1 OR usg.uid IS null)
				AND lo.loginname = system_user
				AND CASE
					WHEN (usg.uid IS null) THEN null
					ELSE usg.name
				END NOT LIKE ''ASRSys%'' AND usg.name NOT LIKE ''db_owner''
		
			IF (@iFound > 0)
			BEGIN
				SELECT	@psUserName = usu.name,
					@psUserGroup = CASE 
						WHEN (usg.uid IS null) THEN null
						ELSE usg.name
					END,
					@piUserGroupID = usg.gid
				FROM sysusers usu 
				LEFT OUTER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
				LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
				WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
					AND (usg.issqlrole = 1 OR usg.uid IS null)
					AND lo.loginname = system_user
					AND CASE 
						WHEN (usg.uid IS null) THEN null
						ELSE usg.name
						END NOT LIKE ''ASRSys%'' AND usg.name NOT LIKE ''db_owner''
			END
			ELSE
			BEGIN
				SELECT @psUserName = usu.name, 
					@psUserGroup = CASE
						WHEN (usg.uid IS null) THEN null
						ELSE usg.name
					END,
					@piUserGroupID = usg.gid
				FROM sysusers usu 
				LEFT OUTER JOIN (sysmembers mem INNER JOIN sysusers usg ON mem.groupuid = usg.uid) ON usu.uid = mem.memberuid
				LEFT OUTER JOIN master.dbo.syslogins lo ON usu.sid = lo.sid
				WHERE (usu.islogin = 1 AND usu.isaliased = 0 AND usu.hasdbaccess = 1) 
					AND (usg.issqlrole = 1 OR usg.uid IS null)
					AND is_member(lo.loginname) = 1
					AND CASE
						WHEN (usg.uid IS null) THEN null
						ELSE usg.name
					END NOT LIKE ''ASRSys%'' AND usg.name NOT LIKE ''db_owner''
			END
		
			IF @psUserGroup <> ''''
			BEGIN
				DELETE FROM [ASRSysUserGroups] 
				WHERE [UserName] = SUSER_NAME()
		
				INSERT INTO [ASRSysUserGroups] 
				VALUES 
				(
					CASE
						WHEN @sSQLVersion <= 8 THEN USER_NAME()
						ELSE SUSER_NAME()
					END,
					@psUserGroup
				)
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
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersInGroups]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersInGroups]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE spASRGetCurrentUsersInGroups 
		(
			@psGroupNames VARCHAR(4000)
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
		
			DECLARE @IN varchar(4000), 
					@INGroup varchar(4000),
					@Pos int
		
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
		
		
		END'

	EXECUTE (@sSPCode_0)
	

/* ------------------------------------------------------------- */
PRINT 'Step 12 of X - Updating Purge Periods'


	/* ASRSysWorkflowElements - Add new DescHasWorkflowName column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysPurgePeriods')
	and name = 'LastPurgeDate'

	if @iRecCount = 0
	BEGIN
        SELECT @NVarCommand = 'ALTER TABLE ASRSysPurgePeriods ADD 
                               LastPurgeDate [datetime] NULL'
		EXEC sp_executesql @NVarCommand
	END

    SELECT @NVarCommand = 'IF NOT EXISTS(SELECT * FROM ASRSysPurgePeriods WHERE PurgeKey = ''EMAIL'')
                           INSERT ASRSYSPurgePeriods (PurgeKey, Period, Unit, LastPurgeDate)
                           VALUES (''EMAIL'', null, null, null)'
    EXEC sp_executesql @NVarCommand

    SELECT @NVarCommand = 'IF NOT EXISTS(SELECT * FROM ASRSysPurgePeriods WHERE PurgeKey = ''WORKFLOW'')
                           INSERT ASRSYSPurgePeriods (PurgeKey, Period, Unit, LastPurgeDate)
                           VALUES (''WORKFLOW'', null, null, null)'
    EXEC sp_executesql @NVarCommand


	----------------------------------------------------------------------
	-- sp_ASRPurgeDate
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRPurgeDate]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRPurgeDate]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRPurgeDate]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[sp_ASRPurgeDate]
		(
		    @purgedate varchar(8000) OUTPUT,
		    @purgekey varchar(8000)
		)
		AS
		BEGIN
		    DECLARE @unit char(1),
		            @period int,
		            @lastPurge datetime,
		            @today datetime
		
		    /* Only get date and not current time */
		    select @today = convert(datetime,convert(varchar,getdate(),101))
		
		    /* Get purge period details */
		    select @unit = unit
		         , @period = (period * -1)
		         , @lastPurge = lastpurgedate
		    from   asrsyspurgeperiods
		    where  purgekey = @purgekey
		
		    /* calculate purge date */
		    SELECT @purgedate = CASE @unit
		        WHEN ''D'' THEN dateadd(dd,@period,@today)
		        WHEN ''W'' THEN dateadd(ww,@period,@today)
		        WHEN ''M'' THEN dateadd(mm,@period,@today)
		        WHEN ''Y'' THEN dateadd(yy,@period,@today)
		    END
		
		    IF @purgedate IS NULL OR datediff(d,@purgedate,@lastPurge) > 0
		    BEGIN
		      SET @purgedate = @lastPurge
		    END
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- sp_ASRPurgeRecords
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRPurgeRecords]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRPurgeRecords]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRPurgeRecords]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[sp_ASRPurgeRecords]
		(
		    @PurgeKey varchar(8000),
		    @TableName varchar(8000),
		    @DateColumn varchar(8000)
		)
		AS
		BEGIN
		
		    /* EXEC sp_ASRPurgeRecords ''EMAIL'', ''ASRSysEmailQueue'', ''DateDue'' */
		
		    DECLARE @PurgeDate datetime
		    DECLARE @sSQL nvarchar(1000)
		
		    EXEC sp_ASRPurgeDate @PurgeDate OUTPUT, @PurgeKey
		
		    SELECT @sSQL = ''DELETE FROM '' + @TableName + '' WHERE '' + @DateColumn + '' < '''''' + convert(varchar,@PurgeDate,101) + ''''''''
		    EXEC sp_executesql @sSQL
		
		    --IF @PurgeKey = ''EMAIL'' OR @PurgeKey = ''WORKFLOW''
		    --    UPDATE ASRSysPurgePeriods SET LastPurgeDate = @PurgeDate
		    --    WHERE PurgeKey = @PurgeKey
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRWorkflowLogPurge
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowLogPurge]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowLogPurge]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowLogPurge]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRWorkflowLogPurge] 
				AS
				BEGIN
						EXEC sp_ASRPurgeRecords ''WORKFLOW'', ''ASRSysWorkflowInstances'', ''completionDateTime''
				END'

	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 13 of X - Send Message store procedure'


	----------------------------------------------------------------------
	-- sp_ASRSendMessage
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRSendMessage]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRSendMessage]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRSendMessage]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[sp_ASRSendMessage] 
		(
		@psMessage	varchar(8000),
		@psSPIDS	varchar(8000)
		)
		AS
		BEGIN
			DECLARE @iDBid	integer,
				@iSPid		integer,
				@iUid		integer,
				@sLoginName	varchar(256),
				@dtLoginTime	datetime, 
				@sCurrentUser	varchar(256),
				@sCurrentApp	varchar(256),
				@Realspid integer
		
			CREATE TABLE #tblCurrentUsers				
				(
					hostname varchar(256)
					,loginame varchar(256)
					,program_name varchar(256)
					,hostprocess varchar(20)
					,sid binary(86)
					,login_time datetime
					,spid int
					,uid smallint
				)
			INSERT INTO #tblCurrentUsers
				EXEC spASRGetCurrentUsers
		
			--MH20040224 Fault 8062
			--{
			--Need to get spid of parent process
			SELECT @Realspid = a.spid
			FROM #tblCurrentUsers a
			FULL OUTER JOIN #tblCurrentUsers b
				ON a.hostname = b.hostname
				AND a.hostprocess = b.hostprocess
				AND a.spid <> b.spid
			WHERE b.spid = @@Spid
		
			--If there is no parent spid then use current spid
			IF @Realspid is null SET @Realspid = @@spid
			--}
		
			/* Get the process information for the current user. */
			SELECT @iDBid = db_id(), 
				@sCurrentUser = loginame,
				@sCurrentApp = program_name
			FROM #tblCurrentUsers
			WHERE spid = @@Spid
		
			/* Get a cursor of the other logged in HR Pro users. */
			DECLARE logins_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT DISTINCT spid, loginame, uid, login_time
				FROM #tblCurrentUsers
				WHERE (spid <> @@spid and spid <> @Realspid)
				AND (@psSPIDS = '''' OR charindex('' ''+convert(varchar,spid)+'' '', @psSPIDS)>0)
		
			OPEN logins_cursor
			FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime
			WHILE (@@fetch_status = 0)
			BEGIN
				/* Create a message record for each HR Pro user. */
				INSERT INTO ASRSysMessages 
					(loginname, message, loginTime, dbid, uid, spid, messageTime, messageFrom, messageSource) 
					VALUES(@sLoginName, @psMessage, @dtLoginTime, @iDBid, @iUid, @iSPid, getdate(), @sCurrentUser, @sCurrentApp)
		
				FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime
			END
			CLOSE logins_cursor
			DEALLOCATE logins_cursor
		
			IF OBJECT_ID(''tempdb..#tblCurrentUsers'', N''U'') IS NOT NULL
				DROP TABLE #tblCurrentUsers
		
		END'

	EXECUTE (@sSPCode_0)

/* ------------------------------------------------------------- */
PRINT 'Step 14 of X - Updating permission icons'

	DELETE FROM ASRSysPermissionCategories WHERE categoryID = 1
	INSERT INTO ASRSysPermissionCategories (categoryID, description, picture, listOrder, categoryKey)
		VALUES(1,'Module Access','',1,'MODULEACCESS')
	SELECT @ptrval = TEXTPTR(picture) FROM ASRSysPermissionCategories WHERE categoryID = 1
	WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x0000010001001010000000000000680500001600000028000000100000002000000001000800000000004001000000000000000000000000000000000000000000001C98F80099417C00FFFFFF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000101010101000000000000000000000001010101010000000000000000000000010101010100000000000000000000000101010101000000000000000000000001010101010000000000000202020202000000000000000000000002020202020000000000000000000000020202020200000000000000000000000202020202000000010101010100000002020202020000000101010101000000000000000000000001010101010000000000000000000000010101010100000000000000000000000101010101000000000000000000000000000000000000FFFF0000FFFF0000FC1F0000FC1F0000FC1F0000FC1F0000FC1F000083FF000083FF000083FF00008383000083830000FF830000FF830000FF830000FFFF000000


/* ------------------------------------------------------------- */
PRINT 'Step 15 of X - Modifying Export Details Table'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysExportDetails')
	and name = 'SuppressNulls'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysExportDetails ADD 
								[SuppressNulls] [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysExportDetails 
								SET [ASRSysExportDetails].[SuppressNulls] = 0'
		EXEC sp_executesql @NVarCommand
	END
	
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
PRINT 'Step X of X - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '3.6')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '3.6.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.4.0')

delete from asrsyssystemsettings
where [Section] = '.NET Assembly' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('.NET Assembly', 'minimum version', '3.5.0')

delete from asrsyssystemsettings
where [Section] = 'outlook service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('outlook service', 'minimum version', '3.6.0')

delete from asrsyssystemsettings
where [Section] = 'workflow service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('workflow service', 'minimum version', '3.6.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v3.6')


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
	GRANT EXECUTE ON xp_LoginConfig TO public
	GRANT EXECUTE ON xp_EnumGroups TO public'
EXEC sp_executesql @NVarCommand

-- Version specific functions
IF (@iSQLVersion < 11)
BEGIN
	SELECT @NVarCommand = 'USE master
		GRANT EXECUTE ON xp_StartMail TO public
		GRANT EXECUTE ON xp_SendMail TO public';
	EXEC sp_executesql @NVarCommand;
END


SELECT @NVarCommand = 'USE ['+@DBName + ']'
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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v3.6 Of HR Pro'
