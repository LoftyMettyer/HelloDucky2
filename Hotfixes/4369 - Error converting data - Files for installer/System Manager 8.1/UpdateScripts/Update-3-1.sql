
/* ----------------------------------------------------- */
/* Update the database from version 3.0 to version 3.1 */
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
	@iSQLVersion numeric(3,1),
	@iTemp integer,
	@sTemp varchar(8000),
	@sTemp2 varchar(8000),
	@sTemp3 varchar(8000)

DECLARE @sGroup sysname
DECLARE @sObject sysname
DECLARE @sObjectType char(2)
DECLARE @sSQL varchar(8000)
DECLARE @sSPCode_0 nvarchar(4000)
DECLARE @sSPCode_1 nvarchar(4000)
DECLARE @sSPCode_2 nvarchar(4000)
DECLARE @sSPCode_3 nvarchar(4000)
DECLARE @sSPCode_4 nvarchar(4000)
/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON
SET @DBName = DB_NAME()
SELECT @iSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY('ProductVersion')));

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */


SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not version 3.0 or 3.1. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '3.0') and (@sDBVersion <> '3.1')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

/* ------------------------------------------------------------- */
PRINT 'Step 1 of X - Modifying columns for Workflow'

	/* Add new record selector identifier columns */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'RecSelWebFormIdentifier'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						[RecSelWebFormIdentifier] [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'RecSelIdentifier'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						[RecSelIdentifier] [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'SecondaryDataRecord'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						[SecondaryDataRecord] [int] NULL'
		EXEC sp_executesql @NVarCommand

	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'SecondaryRecSelWebFormIdentifier'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						[SecondaryRecSelWebFormIdentifier] [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'SecondaryRecSelIdentifier'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						[SecondaryRecSelIdentifier] [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'ForeColorHighlight'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						[ForeColorHighlight] [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'BackColorHighlight'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						[BackColorHighlight] [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementColumns')
	and name = 'DBColumnID'
	
	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementColumns ADD 
						[DBColumnID] [int] NULL'
		EXEC sp_executesql @NVarCommand
	END
	
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementColumns')
	and name = 'DBRecord'
	
	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementColumns ADD 
						[DBRecord] [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'ValueDescription'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						[ValueDescription] [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'EmailSubject'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						[EmailSubject] [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysWorkflowElements
						SET EmailSubject = ''HR Pro Workflow''
						WHERE type = 3'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysEmailQueue')
	and name = 'WorkflowInstanceID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysEmailQueue ADD 
						[WorkflowInstanceID] [int] NULL'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 2 of X - Modifying stored procedures for Workflow'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRRecordDescription]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRRecordDescription]

	SET @sTemp = 'CREATE PROCEDURE spASRRecordDescription
		(
			@piTableID		integer,
			@piRecordID		integer,
			@psRecordDescription	varchar(8000)	OUTPUT
		)
		 AS
		BEGIN
			DECLARE @sSQL varchar(8000),
				@iRecordDescID integer,
				@sRecordDesc varchar(8000)
		
			SET @psRecordDescription = ''''
		
			SELECT @iRecordDescID = ISNULL(ASRSysTables.recordDescExprID, 0)
			FROM ASRSysTables
			WHERE ASRSysTables.tableID = @piTableID
		
			IF @iRecordDescID > 0 
			BEGIN
				SET @sSQL = ''sp_ASRExpr_'' + convert(varchar,@iRecordDescID)
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					EXEC @sSQL @sRecordDesc OUTPUT, @piRecordID
					SET @psRecordDescription = @sRecordDesc
				END
			END
		END'

	EXEC (@sTemp)


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetWorkflowGridItems]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetWorkflowGridItems]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRGetWorkflowGridItems
			(
				@piInstanceID		integer,
				@piElementItemID	integer
			)
			AS
			BEGIN
				DECLARE 
					@iTableID 		integer,
					@iDfltOrderID		integer,
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
					@iDataType		integer,
					@iRecordID		integer,
					@iPersonnelTableID	integer,
					@iWorkflowID		integer,
					@iElementType		integer
			
				SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel''
			
				SELECT 			
					@iTableID = ASRSysWorkflowElementItems.tableID,
					@sRecSelWebFormIdentifier = isnull(ASRSysWorkflowElementItems.wfFormIdentifier, ''''),
					@sRecSelIdentifier = isnull(ASRSysWorkflowElementItems.wfValueIdentifier, 0),
					@iDBRecord = ASRSysWorkflowElementItems.dbRecord,
					@iDfltOrderID = ASRSysTables.defaultOrderID,
					@sBaseTableName = ASRSysTables.tableName
				FROM ASRSysWorkflowElementItems
				INNER JOIN ASRSysTables ON ASRSysWorkflowElementItems.tableID = ASRSysTables.tableID
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iWorkflowID = ASRSysWorkflowInstances.workflowID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				SET @sSelectSQL = ''''
				SET @sOrderSQL = ''''
			
				CREATE TABLE #joinParents
				(
					tableID		integer
				)	
			
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
				WHERE ASRSysOrderItems.orderID = @iDfltOrderID
				ORDER BY ASRSysOrderItems.type,
					ASRSysOrderItems.sequence
			
				OPEN orderCursor
				FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTempTableID, @iTempTableType, @sTempTableName, @sOrderItemType, @fAscending
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @sOrderItemType = ''F''
					BEGIN
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
						FROM #joinParents
						WHERE tableID = @iTempTableID
			
						IF @iTempCount = 0
						BEGIN
							INSERT INTO #joinParents (tableID) VALUES(@iTempTableID)
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
						#joinParents.tableID
					FROM #joinParents
					INNER JOIN ASRSysTables ON #joinParents.tableID = ASRSysTables.tableID
			
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
						SET @sSelectSQL = @sSelectSQL + 
							'' WHERE '' + @sBaseTableName + ''.ID_'' + convert(varchar(100), @iPersonnelTableID) + '' = '' + convert(varchar(100), @iInitiatorID)
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
							SELECT @iRecordID = convert(integer, ISNULL(IV.value, ''0'')),
								 @iTempTableID = EI.tableID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
							INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
							WHERE IV.instanceID = @piInstanceID
								AND IV.identifier = @sRecSelIdentifier
								AND Es.identifier = @sRecSelWebFormIdentifier
								AND Es.workflowID = @iWorkflowID
						END
						ELSE
						BEGIN
							-- StoredData
							SELECT @iRecordID = convert(integer, ISNULL(IV.value, ''0'')),
								 @iTempTableID = Es.dataTableID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sRecSelWebFormIdentifier
							WHERE IV.instanceID = @piInstanceID
						END
			
						SET @sSelectSQL = @sSelectSQL + 
							'' WHERE '' + @sBaseTableName + ''.ID_'' + convert(varchar(100), @iTempTableID) + '' = '' + convert(varchar(100), @iRecordID)
					END
			
					SET @sSelectSQL = @sSelectSQL + 
						'' ORDER BY '' + @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN '','' 
							ELSE '''' 
						END + 
						@sBaseTableName + ''.ID''
			
					EXEC (@sSelectSQL)
				END
			END'

	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetStoredDataActionDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetStoredDataActionDetails]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRGetStoredDataActionDetails
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
					@iSDColumnID int	
						
				SET @psSQL = ''''
				SET @piRecordID = 0
			
				SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel''
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID
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
					@iWorkflowID = workflowID
				FROM ASRSysWorkflowElements
				WHERE ID = @piElementID
			
				SELECT @psTableName = tableName
				FROM ASRSysTables
				WHERE tableID = @piDataTableID
			
				IF @iDataRecord = 0 -- 0 = Initiator''s record
				BEGIN
					SET @piRecordID = @iInitiatorID
			
					IF @piDataTableID = @iPersonnelTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iPersonnelTableID)
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
						SELECT @piRecordID = convert(integer, ISNULL(IV.value, ''0'')),
							 @iTempTableID = EI.tableID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
						INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
						WHERE IV.instanceID = @piInstanceID
							AND IV.identifier = @sRecSelIdentifier
							AND Es.identifier = @sRecSelWebFormIdentifier
							AND Es.workflowID = @iWorkflowID
					END
					ELSE
					BEGIN
						-- StoredData
						SELECT @piRecordID = convert(integer, ISNULL(IV.value, ''0'')),
							 @iTempTableID = Es.dataTableID
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
							AND Es.workflowID = @iWorkflowID
							AND Es.identifier = @sRecSelWebFormIdentifier
						WHERE IV.instanceID = @piInstanceID
					END
				
					IF @piDataTableID = @iTempTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iTempTableID)
					END
				END
			
				IF @piDataAction = 0 -- Insert
				BEGIN
					IF @iSecondaryDataRecord = 0 -- 0 = Initiator''s record
					BEGIN
						SET @iSecondaryRecordID = @iInitiatorID
				
						IF @piDataTableID = @iPersonnelTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iPersonnelTableID)
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
							SELECT @iSecondaryRecordID = convert(integer, ISNULL(IV.value, ''0'')),
								 @iTempTableID = EI.tableID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
							INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
							WHERE IV.instanceID = @piInstanceID
								AND IV.identifier = @sSecondaryRecSelIdentifier
								AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
								AND Es.workflowID = @iWorkflowID
						END
						ELSE
						BEGIN
							-- StoredData
							SELECT @iSecondaryRecordID = convert(integer, ISNULL(IV.value, ''0'')),
								 @iTempTableID = Es.dataTableID
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sSecondaryRecSelWebFormIdentifier
							WHERE IV.instanceID = @piInstanceID
						END
						
						IF @piDataTableID = @iTempTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iTempTableID)
						END
					END
				END

				IF @piDataAction = 0 OR @piDataAction = 1
				BEGIN
					/* INSERT or UPDATE. */
					SET @sColumnList = ''''
					SET @sValueList = ''''

					CREATE TABLE #dbValues (ID integer, 
						wfFormIdentifier varchar(1000),
						wfValueIdentifier varchar(1000),
						dbColumnID int,
						dbRecord int,
						value varchar(8000))

					INSERT INTO #dbValues (ID, 
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
					FROM #dbValues'

	SET @sTemp2 = '
					OPEN dbValuesCursor
					FETCH NEXT FROM dbValuesCursor INTO @iID,
						@sWFFormIdentifier,
						@sWFValueIdentifier,
						@iDBColumnID,
						@iDBRecord
					WHILE (@@fetch_status = 0)
					BEGIN
						SELECT @sDBTableName = tbl.tableName,
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
							-- Initiator record
							SET @iRecordID = @iInitiatorID
						END
						IF @iDBRecord = 1
						BEGIN
							-- Identified record
							SELECT @iRecordID = IV.value
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements WE ON IV.elementID = WE.ID
								AND WE.identifier = @sWFFormIdentifier 
							WHERE IV.instanceID = @piInstanceID
								AND CASE
									WHEN WE.type = 5 THEN 1 -- StoredData
									ELSE  -- WebForm
										CASE 
											WHEN IV.identifier = @sWFValueIdentifier THEN 1
											ELSE 0
										END
									END = 1
						END

						SET @sSQL = @sSQL + convert(nvarchar(4000), @iRecordID)
						SET @sParam = N''@sDBValue varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT

						UPDATE #dbValues
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
							WHEN EC.valueType = 0 THEN EC.value -- Fixed Value
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
							EC.ID
					FROM ASRSysWorkflowElementColumns EC
					INNER JOIN ASRSysColumns SC ON EC.columnID = SC.columnID
					WHERE EC.elementID = @piElementID
			
					OPEN columnCursor
					FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @iValueType = 2 -- DBValue - get here to avoid collation conflict
						BEGIN
							SELECT @sValue = dbV.value
							FROM #dbValues dbV
							WHERE dbV.ID = @iSDColumnID
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
									ELSE isnull(@sValue, 0) -- integer, logic, numeric
								END
						END
			
						FETCH NEXT FROM columnCursor INTO @iColumnID, @sColumnName, @iColumnDataType, @sValue, @iValueType, @iSDColumnID
					END
			
					CLOSE columnCursor
					DEALLOCATE columnCursor
			
					DROP TABLE #dbValues
			
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
									OR @iSecondaryDataRecord = 1) -- 1 = Previous record selector''s record
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
			END'

	EXEC (@sTemp + @sTemp2)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetWorkflowEmailMessage]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	    drop procedure [dbo].[spASRGetWorkflowEmailMessage]

	SET @sSPCode_0 = 'CREATE PROCEDURE dbo.spASRGetWorkflowEmailMessage
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
					@iWorkflowID		integer
							
				SET @psMessage = ''''
			
				exec spASRGetSetting 
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
	
				SELECT @sParam1 = parameterValue
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''		
					AND parameterKey = ''Param_Web1''
				
				SET @sDBName = db_name()
	
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iWorkflowID = ASRSysWorkflowInstances.workflowID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				DECLARE itemCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT EI.caption,
					EI.itemType,
					EI.dbColumnID,
					EI.dbRecord,
					EI.wfFormIdentifier,
					EI.wfValueIdentifier, 
					EI.recSelWebFormIdentifier,
					EI.recSelIdentifier
				FROM ASRSysWorkflowElementItems EI
				WHERE EI.elementID = @piElementID
				ORDER BY EI.ID
			
				OPEN itemCursor
				FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @iItemType = 1
					BEGIN
						/* Database value. */
						SELECT @sTableName = ASRSysTables.tableName, 
							@sColumnName = ASRSysColumns.columnName, 
							@iSourceItemType = ASRSysColumns.dataType
						FROM ASRSysColumns
						INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
						WHERE ASRSysColumns.columnID = @iDBColumnID
			
						IF @iDBRecord = 0 SET @iRecordID = @iInitiatorID
	
						IF @iDBRecord = 1
						BEGIN
							-- Previously identified record.
							SELECT @iElementType = ASRSysWorkflowElements.type
							FROM ASRSysWorkflowElements
							WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
								AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))
	
							IF @iElementType = 2
							BEGIN
								 -- WebForm
								SELECT @iRecordID = convert(integer, ISNULL(IV.value, ''0''))
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
								INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
								WHERE IV.instanceID = @piInstanceID
									AND IV.identifier = @sRecSelIdentifier
									AND Es.identifier = @sRecSelWebFormIdentifier
									AND Es.workflowID = @iWorkflowID
							END
							ELSE
							BEGIN
								-- StoredData
								SELECT @iRecordID = convert(integer, ISNULL(IV.value, ''0''))
								FROM ASRSysWorkflow'



	SET @sSPCode_1 = 'InstanceValues IV
								INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
									AND Es.workflowID = @iWorkflowID
									AND Es.identifier = @sRecSelWebFormIdentifier
								WHERE IV.instanceID = @piInstanceID
							END
						END		
	
						SET @sSQL = ''SELECT @sValue = '' + @sTableName + ''.'' + @sColumnName +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
						SET @sSQLParam = N''@sValue varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT
			
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
							@iSourceItemType = ASRSysWorkflowElementItems.itemType
						FROM ASRSysWorkflowInstanceValues
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceValues.elementID = ASRSysWorkflowElements.ID
						INNER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.elementID
						WHERE ASRSysWorkflowElements.identifier = @sWFFormIdentifier
							AND ASRSysWorkflowInstanceValues.identifier = @sWFValueIdentifier
							AND ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
							AND ASRSysWorkflowElementItems.identifier = @sWFValueIdentifier
			
						IF @sValue IS null SET @sValue = ''''
	
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
	
			
					FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier
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
		'



	SET @sSPCode_2 = '			
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
					SELECT SE.elementID, ISNULL(WE.caption, '''')
					FROM #succeedingElements SE
					INNER JOIN ASRSysWorkflowElements WE ON SE.elementID = WE.ID
				
					OPEN elementCursor
					FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
					WHILE (@@fetch_status = 0)
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
						END
									
						IF LEN(@sQueryString) = 0 
						BEGIN
							SET @psMessage = @psMessage + CHAR(13) +
								@sCaption + '' - Error constructing the query string. Please contact your system administrator.''
						END
						ELSE
						BEGIN
							SET @psMessage = @psMessage + CHAR(13) +
								@sCaption + '' - '' + @sURL + ''/?'' + @sQueryString
						END
						
						FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
					END
					CLOSE elementCursor
			
					DEALLOCATE elementCursor
				END
			
				DROP TABLE #succeedingElements
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetWorkflowFormItems]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	    drop procedure [dbo].[spASRGetWorkflowFormItems]

	SET @sSPCode_0 = 'CREATE PROCEDURE dbo.spASRGetWorkflowFormItems
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
			@iElementType		integer
				
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
					@sColumnName = ASRSysColumns.columnName,
					@iDBColumnDataType = ASRSysColumns.dataType
				FROM ASRSysColumns
				INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
				WHERE ASRSysColumns.columnID = @iDBColumnID
	
				IF @iDBRecord = 0 SET @iRecordID = @iInitiatorID
				IF @iDBRecord = 1
				BEGIN
					-- Identified record.
					SELECT @iElementType = ASRSysWorkflowElements.type
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
						AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sWFFormIdentifier)))
					IF @iElementType = 2
					BEGIN
						 -- WebForm
						SELECT @iRecordID = convert(integer, ISNULL(IV.value, ''0''))
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElement'



	SET @sSPCode_1 = 'Items EI ON IV.identifier = EI.identifier
						INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
						WHERE IV.instanceID = @piInstanceID
							AND IV.identifier = @sWFValueIdentifier
							AND Es.identifier = @sWFFormIdentifier
							AND Es.workflowID = @iWorkflowID
					END
					ELSE
					BEGIN
						-- StoredData
						SELECT @iRecordID = convert(integer, ISNULL(IV.value, ''0''))
						FROM ASRSysWorkflowInstanceValues IV
						INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
							AND Es.workflowID = @iWorkflowID
							AND Es.identifier = @sWFFormIdentifier
						WHERE IV.instanceID = @piInstanceID
					END
				END		
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
	
		SELECT thisFormItems.*, 
			#itemValues.value, 
			CASE
				WHEN thisFormItems.itemType = 4 THEN sourceItems.itemType 
				WHEN thisFormItems.itemType = 1 THEN sourceColumns.dataType 
				ELSE null
			END AS [sourceItemType]
		FROM ASRSysWorkflowElementItems thisFormItems
		LEFT OUTER JOIN #itemValues ON thisFormItems.ID = #itemValues.ID
		LEFT OUTER JOIN ASRSysWorkflowElements sourceElements ON thisFormItems.WFFormIdentifier = sourceElements.identifier
			AND len(isnull(thisFormItems.WFFormIdentifier, '''')) > 0 
			AND sourceElements.workflowID = @iWorkflowID
		LEFT OUTER JOIN ASRSysWorkflowElementItems sourceItems ON sourceElements.id = sourceItems.elementID
			AND thisFormItems.WFValueIdentifier = sourceItems.identifier
		LEFT OUTER JOIN ASRSysColumns sourceColumns ON thisFormItems.DBColumnID = sourceColumns.columnID
			AND thisFormItems.DBColumnID > 0 
		WHERE thisFormItems.elementID = @piElementID
		ORDER BY thisFormItems.ZOrder DESC
		DROP TABLE #itemValues
	END'


	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRCancelPendingPrecedingWorkflowElements]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRCancelPendingPrecedingWorkflowElements]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRCancelPendingPrecedingWorkflowElements
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
					AND (S.status = 0 OR S.status = 1 OR S.status = 2)
			
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
			END'

	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetWorkflowQueryString]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetWorkflowQueryString]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRGetWorkflowQueryString
			(
				@piInstanceID	integer,
				@piElementID	integer,
				@psQueryString	varchar(8000)	output
			)
			AS
			BEGIN
				DECLARE
					@hResult	integer,
					@objectToken	integer,
					@sURL		varchar(8000),
					@sParam1	varchar(8000),
					@sDBName	sysname
			
				SET @psQueryString = ''''
			
				SELECT @sURL = parameterValue
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
					AND parameterKey = ''Param_URL''
			
				SELECT @sParam1 = parameterValue
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
					AND parameterKey = ''Param_Web1''
			
				IF (len(@sURL) > 0)
				BEGIN
					SET @sDBName = db_name()
			
					EXEC @hResult = sp_OACreate ''vbpHRProServer.clsWorkflow'', @objectToken OUTPUT
			
					IF (@hResult = 0) 
					BEGIN
						EXEC @hResult = sp_OAMethod @objectToken, ''GetQueryString'', @psQueryString OUTPUT, @piInstanceID, @piElementID, @sParam1, @@servername, @sDBName
						IF @hResult <> 0
						BEGIN
							SET @psQueryString = ''''
						END
			
						IF len(@psQueryString) > 0
						BEGIN
							SET @psQueryString = @sURL + ''/?'' + @psQueryString
						END
			
						EXEC sp_OADestroy @objectToken
					END
				END
			END'

	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRWorkflowUsesInitiator]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRWorkflowUsesInitiator]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRWorkflowUsesInitiator
			(
				@piWorkflowID		integer,			
				@pfUsesInitiator		bit	OUTPUT
			)
			AS
			BEGIN
				/* Return 1 if the given workflow uses the initiator''s personnel record; else return 0 */
				DECLARE
					@iCount	integer
			
				SET @pfUsesInitiator = 0
			
				/* Initiator''s record used by a Stored Data element action? */
				SELECT @iCount = COUNT(*)
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.type = 5 -- 5 = Stored Data element
					AND (ASRSysWorkflowElements.dataRecord = 0 OR ASRSysWorkflowElements.secondaryDataRecord = 0) -- 0 = Initiator''s record
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
				IF @iCount > 0 SET @pfUsesInitiator = 1
			
				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by a Stored Data element Database Value item? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElementColumns
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementColumns.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowElements.type = 5 -- 5 = Stored Data element
						AND ASRSysWorkflowElementColumns.valueType = 2 -- 2 = Database value
						AND ASRSysWorkflowElementColumns.dbRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END

				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by an Email element address? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.type = 3 -- 3 = Email element
						AND ASRSysWorkflowElements.emailRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END
			
				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by an Email element Database Value item? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElementItems
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowElements.type = 3 -- 3 = Email element
						AND ASRSysWorkflowElementItems.itemType = 1 -- 1 = Database value
						AND ASRSysWorkflowElementItems.dbRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END
			
				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by a Web Form element Database Value? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElementItems
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowElements.type = 2 -- 2 = Web Form element
						AND ASRSysWorkflowElementItems.itemType = 1 -- 1 = Database value
						AND ASRSysWorkflowElementItems.dbRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END
			
				IF @pfUsesInitiator = 0
				BEGIN
					/* Initiator''s record used by a Web Form element Record Selector? */
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowElementItems
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowElements.type = 2 -- 2 = Web Form element
						AND ASRSysWorkflowElementItems.itemType = 11 -- 11 = Record Selector
						AND ASRSysWorkflowElementItems.dbRecord = 0 -- 0 = Initiator''s record
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
					IF @iCount > 0 SET @pfUsesInitiator = 1
				END
			END'

	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRInstantiateWorkflow]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRInstantiateWorkflow]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRInstantiateWorkflow
			(
				@piWorkflowID		integer,			
				@piInstanceID		integer		OUTPUT,
				@psFormElements	varchar(8000)	OUTPUT,
				@psMessage	varchar(8000)	OUTPUT
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
					@hResult		integer,
					@sActualLoginName sysname,
					@sUserGroupName sysname,
					@iUserGroupID integer,
					@fUsesInitiator	bit, 
					@iTemp int,
					@iStartElementID int,
					@superCursor	cursor		
			
				SET @iInitiatorID = 0
				SET @psFormElements = ''''
				SET @psMessage = ''''
			
				EXEC spASRIntGetActualUserDetails
					@sActualLoginName OUTPUT,
					@sUserGroupName OUTPUT,
					@iUserGroupID OUTPUT	
				
				SET @sSQL = ''spASRSysGetCurrentUserRecordID''
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @hResult = 0
			
					EXEC @hResult = @sSQL 
						@iRecordID OUTPUT,
						@iRecordCount OUTPUT
				END
			
				IF NOT @iRecordID IS null SET @iInitiatorID = @iRecordID
				IF @iInitiatorID = 0 
				BEGIN
					/* Unable to determine the initiator''s record ID. Is it needed anyway? */
					EXEC spASRWorkflowUsesInitiator
						@piWorkflowID,
						@fUsesInitiator OUTPUT	
				
					IF @fUsesInitiator = 1
					BEGIN
						IF @iRecordCount = 0
						BEGIN
							/* No records for the initiator. */
							SET @psMessage = ''Unable to locate your personnel record.''
						END
						IF @iRecordCount > 1
						BEGIN
							/* More than one record for the initiator. */
							SET @psMessage = ''You have more than one personnel record.''
						END
				
						RETURN
					END	
				END
			
				/* Create the Workflow Instance record, and remember the ID. */
				INSERT INTO ASRSysWorkflowInstances (workflowID, initiatorID, status, userName)
				VALUES (@piWorkflowID, @iInitiatorID, 0, @sActualLoginName)
							
				SELECT @piInstanceID = MAX(id)
				FROM ASRSysWorkflowInstances
			
				/* Create the Workflow Instance Steps records. 
				Set the first steps'' status to be 1 (pending Workflow Engine action). 
				Set all subsequent steps'' status to be 0 (on hold). */
			
				SELECT @iStartElementID = ASRSysWorkflowElements.ID
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.type = 0 -- Start element
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				
				CREATE TABLE #succeedingElements (elementID integer)
				
				EXEC spASRGetSucceedingWorkflowElements @iStartElementID, @superCursor OUTPUT
				
				FETCH NEXT FROM @superCursor INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
					
					FETCH NEXT FROM @superCursor INTO @iTemp 
				END
				CLOSE @superCursor
				DEALLOCATE @superCursor

				INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime)
				SELECT 
					@piInstanceID, 
					ASRSysWorkflowElements.ID, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN 3
						WHEN ASRSysWorkflowElements.ID IN (SELECT #succeedingElements.elementID
							FROM #succeedingElements) THEN 1
						ELSE 0
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
						WHEN ASRSysWorkflowElements.ID IN (SELECT #succeedingElements.elementID
							FROM #succeedingElements) THEN getdate()
						ELSE null
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
						ELSE null
					END
				FROM ASRSysWorkflowElements 
				WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID
			
				DROP TABLE #succeedingElements

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
						OR ASRSysWorkflowElementItems.itemType = 0)
				UNION
				SELECT  @piInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElements.identifier
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowElements.type = 5
							
				/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
				DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowInstanceSteps.ID,
					ASRSysWorkflowInstanceSteps.elementID
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND ASRSysWorkflowElements.type = 2
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
			
				OPEN formsCursor
				FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
				WHILE (@@fetch_status = 0) 
				BEGIN
					SET @psFormElements = @psFormElements + convert(varchar(8000), @iElementID) + char(9)
			
					/* Change the step''s status to be 2 (pending user input). */
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 2, 
						userName = @sActualLoginName
					WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID
			
					FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
				END
				CLOSE formsCursor
				DEALLOCATE formsCursor
			END'

	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRSubmitWorkflowStep]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	    drop procedure [dbo].[spASRSubmitWorkflowStep]

	SET @sSPCode_0 = 'CREATE PROCEDURE dbo.spASRSubmitWorkflowStep
		(
			@piInstanceID		integer,
			@piElementID		integer,
			@psFormInput1		varchar(8000),
			@psFormInput2		varchar(8000),
			@psFormElements		varchar(8000)	OUTPUT
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
				@sEmailSubject	varchar(200)
			SET @psFormElements = ''''
						
			/* Get the type of the given element */
			SELECT @iElementType = E.type,
				@iEmailID = E.emailID,
				@iEmailRecord = E.emailRecord, 
				@iWorkflowID = E.workflowID,
				@sRecSelIdentifier = E.RecSelIdentifier, 
				@sRecSelWebFormIdentifier = E.RecSelWebFormIdentifier, 
				@iTableID = E.dataTableID,
				@iDataAction = E.dataAction, 
				@sEmailSubject = ISNULL(E.emailSubject, '''')
			FROM ASRSysWorkflowElements E
			WHERE E.ID = @piElementID
			IF @iElementType = 5 -- Stored Data element
			BEGIN
				SET @sValue = @psFormInput1
				SET @sValueDescription = ''''
				SET @sMessage = ''Successfully '' +
					CASE
						WHEN @iDataAction = 0 THEN ''inserted''
						WHEN @iDataAction = 1 THEN ''updated''
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
						EXEC spASRRecordDescription 
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
			END
			ELSE
			BEGIN
				/* Put the submitted form values into the ASRSysWorkflowInstanceValues table. */
				WHILE charindex(CHAR(9), @psFormInput1 + @psFormInput2) > 0
				BEGIN
					SET @iIndex1 = charindex(CHAR(9), @psFormInput1 + @psFormInput2)
					SET @iIndex2 = charindex(CHAR(9), @psFormInput1 + @psFormInput2, @iIndex1+1)
		
					SET @sID = replace(LEFT(@psFormInput1 + @psFormInput2, @iIndex1-1), '''''''', '''''''''''')
					SET @sValue = SUBSTRING(@psFormInput1 + @psFormInput2, @iIndex1+1, @iIndex2-@iIndex1-1)
					--Get the record description (for RecordSelectors only)
					SET @sValueDescription = ''''
					-- Get the WebForm item type, etc.
					SELECT @sIdentifier = EI.identifier,
						@iItemType = EI.itemType,
						@iTableID = EI.tableID
					FROM ASRSysWorkflowElementItems EI
					WHERE EI.ID = convert(integer, @sID)
					IF @iItemType = 11 -- Record Selector
					BEGIN
						-- Get the table record description ID. 
						SELECT @'



	SET @sSPCode_1 = 'iRecDescID =  ASRSysTables.RecordDescExprID
						FROM ASRSysTables 
						WHERE ASRSysTables.tableID = @iTableID
						-- Get the record description. 
						IF (NOT @iRecDescID IS null) AND (@iRecDescID > 0) AND (convert(integer, @sValue) > 0)
						BEGIN
							SET @iTemp = convert(integer, @sValue)
							SET @sExecString = ''exec sp_ASRExpr_'' + convert(nvarchar(4000), @iRecDescID) + '' @recDesc OUTPUT, @recID''
							SET @sParamDefinition = N''@recDesc varchar(8000) OUTPUT, @recID integer''
							EXEC sp_executesql @sExecString, @sParamDefinition, @sEvalRecDesc OUTPUT, @iTemp
							IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sValueDescription = @sEvalRecDesc
						END
					END
					UPDATE ASRSysWorkflowInstanceValues
					SET ASRSysWorkflowInstanceValues.value = @sValue, 
						ASRSysWorkflowInstanceValues.valueDescription = @sValueDescription
					WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceValues.elementID = @piElementID
						AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier
		
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
			END
					
			SET @hResult = 0
			SET @sTo = ''''
		
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
						/* Initiator record. */
						SELECT @iEmailRecordID = ASRSysWorkflowInstances.initiatorID
						FROM ASRSysWorkflowInstances
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID
					END
		
					IF @iEmailRecord = 1
					BEGIN
						SELECT @iPrevElementType = ASRSysWorkflowElements.type
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))
						IF @iPrevElementType = 2
						BEGIN
							 -- WebForm
							SELECT @iEmailRecordID = convert(integer, ISNULL(IV.value, ''0''))
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
							INNER JOIN ASRSysWorkflowElements Es ON EI.elementID = Es.ID
							WHERE IV.instanceID = @piInstanceID
								AND IV.identifier = @sRecSelIdentifier
								AND Es.identifier = @sRecSelWebFormIdentifier
								AND Es.workflowID = @iWorkflowID
						END
						ELSE
						BEGIN
							-- StoredData
							SELECT @iEmailRecordID = convert(integer, ISNULL(IV.value, ''0''))
							FROM ASRSysWorkflowInstanceValues IV
							INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sRecSelWebFormIdentifier
							WHERE IV.instanceID = @piInstanceID
						END
					END
					/* Get the recipient address. */
					EXEC @hResult = @sSQL @sTo OUTPUT, @iEmailID, @iEmailRecordID
					IF @sTo IS null SET @sTo = ''''
				END
		
				IF LEN(rtrim(ltrim(@sTo))) = 0
				BEGIN
					/* Email step failure if no known recipient.*/
					/* Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. */
					EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''No email recipient.''
					
					SET @hResult = -1
				END
				ELSE
				BEGIN
					/* Build the email message. */
					EXEC spASRGetWorkflowEmailMessage @piInstanceID, @piElementID, @sMessage OUTPUT
		
					/* Send the email. */
					'



	SET @sSPCode_2 = 'INSERT ASRSysEmailQueue(
						RecordDesc,
						ColumnValue, 
						DateDue, 
						UserName, 
						[Immediate],
						RecalculateRecordDesc, 
						RepTo,
						MsgText,
						WorkflowInstanceID, 
						Subject)
					VALUES ('''',
						'''',
						getdate(),
						''HR Pro Workflow'',
						1,
						0, 
						@sTo,
						@sMessage,
						@piInstanceID,
						@sEmailSubject)
				END
			END
		
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
					ASRSysWorkflowInstanceSteps.message = CASE
						WHEN @iElementType = 3 THEN LEFT(@sMessage, 8000)
						WHEN @iElementType = 5 THEN LEFT(@sMessage, 8000)
						ELSE ''''
					END
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
				IF @iElementType = 4 -- Decision element
				BEGIN
					SET @iPrevElementType = 4 -- Decision element
					SET @iPreviousElementID = @piElementID
					WHILE (@iPrevElementType = 4)
					BEGIN
						/* Get the ID of the elements that precede the Decision element. */
						CREATE TABLE #precedingElements (elementID integer)
		
						EXEC spASRGetPrecedingWorkflowElements @iPreviousElementID, @superCursor OUTPUT
			
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
					
						SELECT @iPrevElementType = ASRSysWorkflowElements.type
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.ID = @iPreviousElementID
					END
					
					SELECT @iValue = convert(integer, IV.value)
					FROM ASRSysWorkflowInstanceValues IV
					INNER JOIN ASRSysWorkflowElements E ON IV.identifier = E.trueFlowIdentifier
					WHERE IV.elementID = @iPreviousElementID
						AND IV.instanceid = @piInstanceID
						AND E.ID = @piElementID
				
					IF @iValue IS null SET @iValue = 0
		
					UPDATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.decisionFlow = @iValue
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND ASRSysWorkflowInstanceSteps.elementID = @piElementID
			
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
						AND ASRSysWorkflowInstanceSteps.status = 0
		
					DROP TABLE #succeedingElements2
				END
				ELSE
				BEGIN
					CREATE TABLE #succeedingElements (elementID integer)
		
					EXEC spASRGetSucceedingWorkflowElements @piElementID, @superCursor OUTPUT
		
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEG'



	SET @sSPCode_3 = 'IN
						INSERT INTO #succeedingElements (elementID) VALUES (@iTemp)
						
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
							AND WIS.elementID = @piElementID
						/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
						DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT ASRSysWorkflowInstanceSteps.ID,
							ASRSysWorkflowInstanceSteps.elementID
						FROM ASRSysWorkflowInstanceSteps
						INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT #succeedingElements.elementID
								FROM #succeedingElements)
							AND ASRSysWorkflowElements.type = 2
							AND ASRSysWorkflowInstanceSteps.status = 0
						OPEN formsCursor
						FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
						WHILE (@@fetch_status = 0) 
						BEGIN
							SET @psFormElements = @psFormElements + convert(varchar(8000), @iElementID) + char(9)
							
							/* Change the step status to be 2 (pending user input). */
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 2, 
								ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.userName = @sUserName,
								ASRSysWorkflowInstanceSteps.userEmail = @sUserEmail 
							WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID
								AND ASRSysWorkflowInstanceSteps.status = 0
							
							FETCH NEXT FROM formsCursor INTO @iStepID, @iElementID
						END
						CLOSE formsCursor
						DEALLOCATE formsCursor
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate()
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT #succeedingElements.elementID
								FROM #succeedingElements)
							AND ASRSysWorkflowInstanceSteps.elementID NOT IN 
								(SELECT ASRSysWorkflowElements.ID
								FROM ASRSysWorkflowElements
								WHERE ASRSysWorkflowElements.type = 2)
							AND ASRSysWorkflowInstanceSteps.status = 0
					END
					ELSE
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
							ASRSysWorkflowInstanceSteps.userEmail = CASE
								WHEN (SELECT ASRSysWorkflowElements.type 
									FROM ASRSysWorkflowElements 
									WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @sTo -- 2 = Web Form element
								ELSE ASRSysWorkflowInstanceSteps.userEmail
							END
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT #succeedingElements.elementID
								FROM #succeedingElements)
							AND ASRSysWorkflowInstanceSteps.status = 0
					END
					
					DROP TABLE #succeedingElements
				END
			
				/* Set activated Web Forms to be ''pending'' (to be done by the user) */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 2
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHE'



	SET @sSPCode_4 = 'RE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 2)
		
				/* Set activated Terminators to be ''completed'' */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate()
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND ASRSysWorkflowElements.type = 1)
		
				/* Count how many terminators have completed. ie. if the workflow has completed. */
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
					
					/* NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					is performed by a DELETE trigger on the ASRSysWorkflowInstances table. */
				END
				IF @iElementType = 3 -- Email element
					OR @iElementType = 5 -- Stored Data element
				BEGIN
					exec spASREmailImmediate ''HR Pro Workflow''
				END
			END
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRActionActiveWorkflowSteps]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRActionActiveWorkflowSteps]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRActionActiveWorkflowSteps
		AS
		BEGIN
			/* Return a recordset of the workflow steps that need to be actioned by the Workflow service.
			Action any that can be actioned immediately. */
			DECLARE
				@iAction			integer, -- 0 = do nothing, 1 = submit step, 2 = change status to ''2'', 3 = Summing Junction check, 4 = Or check
				@iElementType		integer,
				@iInstanceID		integer,
				@iElementID			integer,
				@iStepID			integer,
				@iCount				integer,
				@sStatus			bit,
				@sMessage			varchar(8000),
				@superCursor		cursor,
				@superCursor2		cursor,
				@iTemp				integer, 
				@iTemp2				integer, 
				@sForms 			varchar(8000), 
				@iType				integer,
				@iDecisionFlow		integer,
				@iInvalidDecisionCount	integer
		
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
					SET @iInvalidDecisionCount = 0
		
					CREATE TABLE #precedingElements (elementID integer)
				
					EXEC spASRGetPrecedingWorkflowElements @iElementID, @superCursor OUTPUT
			
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
						INSERT INTO #precedingElements (elementID) VALUES (@iTemp)
		
						/* Check that the preceding element, if it was a Decision element,
						not only completed, but also followed onto the path to the current element. */
						IF (@iInvalidDecisionCount = 0) 
						BEGIN
							SELECT @iType = WE.type,
								@iDecisionFlow = WIS.decisionFlow
							FROM ASRSysWorkflowInstanceSteps WIS
							INNER JOIN ASRSysWorkflowElements WE ON WIS.elementID = WE.ID
							WHERE WIS.instanceID = @iInstanceID
								AND WE.ID = @iTemp
		
							IF (@iType = 4) -- Decision
							BEGIN
								CREATE TABLE #succeedingElements2 (elementID integer)
								
								EXEC spASRGetDecisionSucceedingWorkflowElements @iTemp, @iDecisionFlow, @superCursor2 OUTPUT
								
								FETCH NEXT FROM @superCursor2 INTO @iTemp2
								WHILE (@@fetch_status = 0)
								BEGIN
									INSERT INTO #succeedingElements2 (elementID) VALUES (@iTemp2)
									
									FETCH NEXT FROM @superCursor2 INTO @iTemp2
								END
								CLOSE @superCursor2
								DEALLOCATE @superCursor2
								
								SELECT @iCount = COUNT(*)
								FROM #succeedingElements2
								WHERE elementID = @iElementID
		
								IF @iCount = 0 SET @iInvalidDecisionCount = @iInvalidDecisionCount + 1
								
								DROP TABLE #succeedingElements2
							END
						END
						
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
					IF (@iCount = 0) AND (@iInvalidDecisionCount = 0) SET @iAction = 1
		
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
					EXEC spASRSubmitWorkflowStep @iInstanceID, @iElementID, '''', '''', @sForms OUTPUT
				END
		
				IF @iAction = 2
				BEGIN
					UPDATE ASRSysWorkflowInstanceSteps
					SET status = 2
					WHERE id = @iStepID
				END
		
				FETCH NEXT FROM stepsCursor INTO @iElementType, @iInstanceID, @iElementID, @iStepID
			END
		END'

	EXEC (@sTemp)


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetActiveWorkflowStoredDataSteps]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetActiveWorkflowStoredDataSteps]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRGetActiveWorkflowStoredDataSteps
			AS
			BEGIN
				/* Return a recordset of the workflow StoredData steps that need to be actioned by the Workflow service. */
			
				CREATE TABLE #Steps (ID integer)
			
				INSERT INTO #Steps
				SELECT S.ID
				FROM ASRSysWorkflowInstanceSteps S
				INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
				WHERE S.status = 1
					AND E.type = 5 -- 5 = Stored Data
			
				UPDATE ASRSysWorkflowInstanceSteps
				SET status = 5 -- In progress
				WHERE ID IN (SELECT ID FROM #Steps)
			
				SELECT S.instanceID AS [instanceID],
					E.ID AS [elementID],
					S.ID AS [stepID]
				FROM ASRSysWorkflowInstanceSteps S
				INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
				WHERE s.ID IN (SELECT ID FROM #Steps)
			
				DROP TABLE #Steps
			END'
	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRWorkflowActionFailed]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRWorkflowActionFailed]

	SET @sTemp = 'CREATE PROCEDURE dbo.spASRWorkflowActionFailed
			(
				@piInstanceID		integer,
				@piElementID		integer,
				@psMessage			varchar(8000)
			)
			AS
			BEGIN
				UPDATE ASRSysWorkflowInstanceSteps
				SET status = 4,	-- 4 = failed
					message = @psMessage
				WHERE instanceID = @piInstanceID
					AND elementID = @piElementID
			
				UPDATE ASRSysWorkflowInstances
				SET status = 2	-- 2 = error
				WHERE ID = @piInstanceID
			END'
	EXEC (@sTemp)

/* ------------------------------------------------------------- */
PRINT 'Step 3 of X - Adding Workflow System Permissions'

	DELETE FROM ASRSysPermissionItems WHERE itemid in (151)
	INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
				VALUES (151,'View log',7,42,'VIEWLOG')

	DELETE FROM ASRSysPermissionItems WHERE itemid in (152)
	INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
				VALUES (152,'Administer',3,42,'ADMINISTER')

	-- Give security to admistrators
	SELECT @iRecCount = count(*)
	FROM ASRSysGroupPermissions
	WHERE itemid IN (151, 152)

	IF @iRecCount = 0 
	BEGIN
		INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
			SELECT DISTINCT 151, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))

		INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
			SELECT DISTINCT 152, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 1 AND permitted = 1) OR (itemid = 3 AND permitted = 1))
	END

/* ------------------------------------------------------------- */
PRINT 'Step 4 of X - Modifying miscellaneous stored procedures'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetCurrentUsers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGetCurrentUsers]

	SET @sTemp = 'CREATE PROCEDURE spASRGetCurrentUsers
			AS
			BEGIN
				SET NOCOUNT ON
			
				SELECT DISTINCT hostname, loginame, program_name, hostprocess
			    FROM ASRTempSysProcesses
			    WHERE program_name like ''HR Pro%'' 
				AND program_name NOT LIKE ''HR Pro Workflow%''
			    AND dbid in ( 
			                   SELECT dbid FROM master..sysdatabases
			                   WHERE name = DB_NAME())
			     ORDER BY loginame
		
			END'
	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRIntCheckUserSessions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRIntCheckUserSessions]

	SET @sTemp = 'CREATE PROCEDURE sp_ASRIntCheckUserSessions 
			(
				@psUserName		varchar(8000),
				@piCount		integer		OUTPUT
			)
			AS
			BEGIN
				SELECT @piCount = COUNT(*)
				FROM master..sysprocesses 
				WHERE program_name like ''HR Pro%'' 
					AND program_name NOT LIKE ''HR Pro Workflow%''
					AND loginame = @psUserName
			END'
	EXEC (@sTemp)

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRSendMessage]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRSendMessage]

	SET @sTemp = 'CREATE PROCEDURE sp_ASRSendMessage 
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
	
				--MH20040224 Fault 8062
				--{
				--Need to get spid of parent process
				SELECT @Realspid = a.spid
				FROM master..sysprocesses a
				FULL OUTER JOIN master..sysprocesses b
					ON a.hostname = b.hostname
					AND a.hostprocess = b.hostprocess
					AND a.spid <> b.spid
				WHERE b.spid = @@Spid
	
				--If there is no parent spid then use current spid
				IF @Realspid is null SET @Realspid = @@spid
				--}
	
	
				/* Get the process information for the current user. */
				SELECT @iDBid = dbid, 
					@sCurrentUser = loginame,
					@sCurrentApp = program_name
				FROM master..sysprocesses
				WHERE spid = @@spid
	
				/* Get a cursor of the other logged in HR Pro users. */
				DECLARE logins_cursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT DISTINCT spid, loginame, uid, login_time
					FROM master..sysprocesses
					WHERE program_name LIKE ''HR Pro%''
					AND program_name NOT LIKE ''HR Pro Workflow%''
					AND dbid = @iDBid
					AND (spid <> @@spid and spid <> @Realspid)
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
			END'
	EXEC (@sTemp)


/* ------------------------------------------------------------- */
PRINT 'Step 5 of X - Modifying absence breakdown stored procedure'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASR_AbsenceBreakdown_Run]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASR_AbsenceBreakdown_Run]

	SET @sTemp = 'CREATE PROCEDURE sp_ASR_AbsenceBreakdown_Run
		(
		@pdReportStart      datetime,
		@pdReportEnd    datetime,
		@pcReportTableName  char(30)
		) 
		AS 
		begin
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

		declare @sSQL as char(8000)
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

			'

	SET @sTemp2 = '/* Calculate total duraton of absence */
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

		END'

	EXEC (@sTemp+@sTemp2)


/* ------------------------------------------------------------- */
PRINT 'Step 6 of X - Modifying audit stored procedures'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRAudit]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRAudit]

	SET @sTemp = 'CREATE PROCEDURE sp_ASRAudit 
		(
			@piColumnID int,
			@piRecordID int,
			@psRecordDesc varchar(255),
			@psOldValue varchar(255),
			@psNewValue varchar(255)
		)
		AS
		BEGIN	
			DECLARE @sTableName varchar(8000),
				@sColumnName varchar(8000)
		
			/* Get the table name for the given column. */
			/* Get the column name for the given column. */
			SELECT @sTableName = ASRSysTables.tableName,
				@sColumnName = ASRSysColumns.columnName
			FROM ASRSysColumns
			INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
			WHERE ASRSysColumns.columnID = @piColumnID
		
			IF @sTableName IS NULL SET @sTableName = ''<Unknown>''
		
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
				(CASE
					WHEN UPPER(LEFT(APP_NAME(), 15)) = ''HR PRO WORKFLOW'' THEN ''HR Pro Workflow''
					ELSE user
				END, 
				getDate(), 
				@sTableName, 
				@piRecordID, 
				@psRecordDesc, 
				@sColumnName, 
				@psOldValue, 
				@psNewValue,
				@piColumnID,
				0)
		END'

	EXEC (@sTemp)


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRAuditTable]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[sp_ASRAuditTable]

	SET @sTemp = 'CREATE PROCEDURE sp_ASRAuditTable 
		(
			@piTableID int,
			@piRecordID int,
			@psRecordDesc varchar(255),
			@psValue varchar(255)
		)
		AS
		BEGIN	
			DECLARE @sTableName varchar(8000)
		
			/* Get the table name for the given column. */
			SELECT @sTableName = tableName 
			FROM ASRSysTables
			WHERE ASRSysTables.tableID = @piTableID
		
			IF @sTableName IS NULL SET @sTableName = ''<Unknown>''
		
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
				0)
		END'

	EXEC (@sTemp)


/* ------------------------------------------------------------- */
PRINT 'Step 7 of X - Updating Accord integration'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysAccordTransferFieldDefinitions')
	and name = 'PreventModify'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransferFieldDefinitions ADD [PreventModify] [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET PreventModify = 0'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions SET PreventModify = 1 WHERE TransferTypeID = 0 AND TransferFieldID = 1'
		EXEC sp_executesql @NVarCommand

	END


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordPopulateTransactionData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordPopulateTransactionData]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordNeedToSendAll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordNeedToSendAll]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordPopulateTransaction]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordPopulateTransaction]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordIsRecordInPayroll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
    drop procedure [dbo].[spASRAccordIsRecordInPayroll]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordDeleteTransactionsForRecord]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
    drop procedure [dbo].[spASRAccordDeleteTransactionsForRecord]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRAccordPopulateTransactionData] (
		@piTransactionID int,
		@piColumnID int,
		@psOldValue varchar(255),
		@psNewValue varchar(255)
		)
	AS
	BEGIN	
		DECLARE @iRecCount int

		SET NOCOUNT ON

		SELECT @iRecCount = COUNT(FieldID) FROM ASRSysAccordTransactionData WHERE @piTransactionID = TransactionID and FieldID = @piColumnID

		-- Insert a record into the Accord Transaction table.	
		IF @iRecCount = 0
			INSERT INTO ASRSysAccordTransactionData
				([TransactionID],[FieldID], [OldData], [NewData])
			VALUES 
				(@piTransactionID,@piColumnID,@psOldValue,@psNewValue)
		ELSE
			UPDATE ASRSysAccordTransactionData SET [OldData] = @psOldValue
				WHERE @piTransactionID = TransactionID and FieldID = @piColumnID
	END'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRAccordNeedToSendAll] 
	(@iRecordID int,
	@bResend bit OUTPUT)
	AS
	BEGIN
		SET NOCOUNT ON

		DECLARE @Status integer

		SELECT TOP 1 @Status = Status FROM ASRSysAccordTransactions
		WHERE HRProRecordID = @iRecordID
		ORDER BY CreatedDateTime DESC
	
		-- Nothing found
		IF @Status IS NULL SET @bResend = 1
	
		-- Previous transaction failed
		IF @Status IN (20) SET @bResend = 0
	
		-- Pending, success, or success with warnings, blocked
		IF @Status IN (1, 10, 11, 21, 22, 30) SET @bResend = 0


	END'
	EXEC sp_executesql @NVarCommand

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRAccordPopulateTransaction] (
	@piTransactionID int OUTPUT,
	@piTransferType int,
	@piTransactionType int ,
	@piDefaultStatus int,
	@piHRProRecordID int,
	@iTriggerLevel int,
	@pbSendAllFields bit OUTPUT)
	AS
	BEGIN	

	-- Return the required user or system setting.
	DECLARE @iCount	integer
	DECLARE @bNewTransaction bit
	DECLARE @iStatus integer
	DECLARE @bCreate bit

	SET NOCOUNT ON

	SET @piTransactionID = null
	SET @bCreate = 1

	SELECT @piTransactionID = TransactionID
		FROM ASRSysAccordTransactionProcessInfo
		WHERE spid = @@SPID AND TransferType = @piTransferType AND RecordID = @piHRProRecordID

	-- Could be a null if the trigger was fired from a non Accord module enabled table, e.g. a child updating a parent field
	IF @piTransactionID IS null SET @bNewTransaction = 1
	ELSE SET @bNewTransaction = 0

	-- Get a transaction ID for this process and update the temporary Accord table
	IF @bNewTransaction = 1
	BEGIN
		SELECT @iCount = COUNT(*)
			FROM ASRSysSystemSettings
			WHERE section = ''AccordTransfer'' AND settingKey = ''NextTransactionID''
		
		IF @iCount = 0
			INSERT ASRSysSystemSettings (Section, SettingKey, SettingValue) VALUES (''AccordTransfer'',''NextTransactionID'',1)
		ELSE
			UPDATE ASRSysSystemSettings SET SettingValue = SettingValue + 1 WHERE section = ''AccordTransfer'' AND settingKey =  ''NextTransactionID''

		SELECT @piTransactionID = settingValue 
		FROM ASRSysSystemSettings
		WHERE section = ''AccordTransfer'' AND settingKey =  ''NextTransactionID''

		-- If update, has it already been sent?
		IF @piTransactionType = 1
		BEGIN
			SELECT TOP 1 @iStatus = Status FROM ASRSysAccordTransactions
			WHERE HRProRecordID = @piHRProRecordID
			ORDER BY CreatedDateTime DESC
		
			IF @iStatus IS NULL OR @iStatus = 20
			BEGIN
				SET @piTransactionType = 0
				SET @pbSendAllFields = 1
			END
		END

		-- Are we trying to delete something thats never been sent?
		IF @piTransactionType = 2
		BEGIN
			SELECT TOP 1 @iStatus = Status FROM ASRSysAccordTransactions
			WHERE HRProRecordID = @piHRProRecordID
			ORDER BY CreatedDateTime DESC
		
			IF @iStatus IS NULL	SET @bCreate = 0
			ELSE SET @pbSendAllFields = 1
		END

		-- Insert a record into the Accord Transfer table.
		IF @bCreate = 1
		BEGIN
			INSERT INTO ASRSysAccordTransactions
				([TransactionID],[TransferType], [TransactionType], [CreatedUser], [CreatedDateTime], [Status], [HRProRecordID], [Archived])
			VALUES 
				(@piTransactionID, @piTransferType, @piTransactionType, SYSTEM_USER, GETDATE(), @piDefaultStatus, @piHRProRecordID, 0)

			INSERT ASRSysAccordTransactionProcessInfo (SPID, TransactionID,TransferType,RecordID) VALUES (@@SPID, @piTransactionID, @piTransferType, @piHRProRecordID)
		END

	END
	END'
	EXEC sp_executesql @NVarCommand


	SET @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRAccordIsRecordInPayroll]
	(@iRecordID int
	,@ProhibitDelete int OUTPUT)
AS
	BEGIN
	SET NOCOUNT ON
	
	DECLARE @cursStatus cursor
	DECLARE @Status integer
	
	SET @ProhibitDelete = 0

	SET @cursStatus = CURSOR LOCAL FAST_FORWARD FOR	
		SELECT Status FROM ASRSysAccordTransactions
		WHERE HRProRecordID = @iRecordID
		ORDER BY CreatedDateTime DESC
	OPEN @cursStatus
	FETCH NEXT FROM @cursStatus INTO @Status
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @Status IN (10,11) SET @ProhibitDelete = 1
		FETCH NEXT FROM @cursStatus INTO @Status
	END	
END	'
	EXEC sp_executesql @NVarCommand

	SET @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRAccordDeleteTransactionsForRecord]
	(@iRecordID int)
	AS
	BEGIN
		SET NOCOUNT ON
		DELETE FROM ASRSysAccordTransactions WHERE HrProRecordID = @iRecordID
	END'
	EXEC sp_executesql @NVarCommand

	-- Add default module status
	SELECT @iRecCount = COUNT(parameterValue) FROM ASRSysModuleSetup WHERE moduleKey = 'MODULE_ACCORD' AND parameterKey = 'Param_AllowStatusChange'
	IF @iRecCount = 0 INSERT INTO ASRSysModuleSetup (ModuleKey, ParameterKey, ParameterValue,ParameterType) VALUES ('MODULE_ACCORD','Param_AllowStatusChange',0,'PType_Other')

	-- Change text of group permission
	UPDATE ASRSysPermissionItems SET Description = 'Administer Transfers' WHERE ItemID = 146


/* ------------------------------------------------------------- */
PRINT 'Step 8 of X - Modifying email stored procedures'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASREmailQueue]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASREmailQueue]

	SET @sTemp = 'CREATE PROCEDURE spASREmailQueue AS
		BEGIN
			DECLARE @sSQL varchar(8000),
				@iQueueID int,
				@iRecordID int,
				@iRecordDescID int,
				@sRecordDesc varchar(8000)
		
			DECLARE emailQueue_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysEmailQueue.queueID, 
				ASRSysEmailQueue.recordID, 
				ASRSysTables.recordDescExprID
			FROM ASRSysEmailQueue
			INNER JOIN ASRSysEmailLinks ON ASRSysEmailQueue.LinkID = ASRSysEmailLinks.LinkID
			INNER JOIN ASRSysColumns ON ASRSysColumns.ColumnID = ASRSysEmailLinks.ColumnID
			INNER JOIN ASRSysTables ON ASRSysTables.TableID = ASRSysColumns.TableID
			WHERE ASRSysEmailQueue.recalculateRecordDesc = 1
			
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
			DEALLOCATE emailQueue_cursor
		END'

	EXEC (@sTemp)


	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASREmailImmediate]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	    drop procedure [dbo].[spASREmailImmediate]


	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].spASREmailImmediate(@Username varchar(255)) AS
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
				LEFT OUTER JOIN ASRSysEmailLinks ON ASRSysEmailLinks.LinkID = ASRSysEmailQueue.LinkID
				WHERE DateSent IS Null And datediff(dd,DateDue,getdate()) >= 0
				AND (LOWER(@Username) = LOWER([Username]) OR @Username = '''')
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
				IF @ColumnID IS null AND @TableID IS null
				BEGIN
					SELECT @emailDate = getDate()
					SELECT @RecipTo = RepTo,
						@CC = RepCC,
						@BCC = RepBCC,
						@Subject = Subject,
						@Attachment = Attachment,
						@MsgText = MsgText
					FROM ASRSysEmailQueue 
					WHERE QueueID = @QueueID
					IF RTrim(@RecipTo) = ''''
						SET @hResult = 1
					ELSE
						EXEC spASRSendMail @hResult OUTPUT, @RecipTo, '''', '''', @Subject,  @MsgText, ''''
				END
				IF @hResult = 0
				BEGIN
					UPDATE ASRSysEmailQueue SET DateSent = @emailDate, RepTo = @RecipTo, RepCC = @CC, RepBCC = @BCC, Subject = @Subject, Attachment = @Attachment
					WHERE QueueID = @QueueID
					
					UPDATE ASRSysEmailQueue SET MsgText = @MsgText
					WHERE QueueID = @QueueID
				END
				FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc, @RecalculateR'

	SET @sSPCode_1 = 'ecordDesc, @TableID, @DateDue
			END
			CLOSE emailqueue_cursor
			DEALLOCATE emailqueue_cursor
		END'


	EXECUTE (@sSPCode_0
		+ @sSPCode_1)



/* ------------------------------------------------------------- */
PRINT 'Step 9 of X - Modifying trigger to Workflow Instances'

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

				DELETE FROM ASRSysEmailQueue
				WHERE ASRSysEmailQueue.workflowInstanceID IN (SELECT id FROM deleted)
			END')


/* ------------------------------------------------------------- */
PRINT 'Step 9 of X - Modifying process checking'

	IF EXISTS (SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[ASRTempSysProcesses]') and OBJECTPROPERTY(id, N'IsTable') = 1)
	DROP TABLE [dbo].[ASRTempSysProcesses]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGenerateSysProcesses]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRGenerateSysProcesses]

	SELECT @NVarCommand = 'CREATE Procedure spASRGenerateSysProcesses
	AS
	BEGIN
		DECLARE @strDBName nvarchar(100)
		DECLARE @objectToken int
		DECLARE @hResult int
		DECLARE @hMessage varchar(255)
		DECLARE @sSQLVersion char(2)
	
		SET @strDBName = DB_Name()
		SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)

		IF @sSQLVersion = ''9''
		BEGIN
	
			  -- Create Server DLL object
			EXEC @hResult = sp_OACreate ''vbpHRProServer.clsSQLFunctions'', @objectToken OUTPUT
			IF @hResult <> 0
			BEGIN
				DELETE FROM [dbo].[ASRTempSysProcesses]
				INSERT INTO [dbo].[ASRTempSysProcesses] SELECT * FROM master..sysprocesses
			END	
			ELSE
			BEGIN
				EXEC @hResult = sp_OAMethod @objectToken, ''GenerateProcesses'', @hMessage OUTPUT, @@SERVERNAME, @strDBName
			END
		END
		ELSE
		BEGIN
			DELETE FROM [dbo].[ASRTempSysProcesses]
			INSERT INTO [dbo].[ASRTempSysProcesses] SELECT * FROM master..sysprocesses
		END

	END'

	EXEC sp_executesql @NVarCommand

	SET @NVarCommand = 'GRANT EXECUTE ON spASRGenerateSysProcesses TO [ASRSysGroup]'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 10 of X - Bypass trigger options'

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysGlobalFunctions')
	and name = 'BypassTrigger'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysGlobalFunctions ADD [BypassTrigger] [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysGlobalFunctions SET BypassTrigger = 0'
		EXEC sp_executesql @NVarCommand
	END


	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysImportName')
	and name = 'BypassTrigger'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysImportName ADD [BypassTrigger] [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = 'UPDATE ASRSysImportName SET BypassTrigger = 0'
		EXEC sp_executesql @NVarCommand
	END



/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step X of X - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '3.1')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '3.1.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.0.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v3.1')


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


---Just in case we have moved SQL versions...
---(Ref 11375-11379 inclusive)
IF EXISTS (SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[ASRTempSysProcesses]') and OBJECTPROPERTY(id, N'IsTable') = 1)
DROP TABLE [dbo].[ASRTempSysProcesses]
SELECT * INTO [dbo].[ASRTempSysProcesses] FROM master..sysprocesses


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v3.1 Of HR Pro'
