
/* --------------------------------------------------- */
/* Update the database from version 3.6 to version 3.7 */
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
IF (@sDBVersion <> '3.6') and (@sDBVersion <> '3.7')
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

	/* ASRSysWorkflows - Add new QueryString column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflows', 'U')
	AND name = 'queryString'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflows ADD 
							queryString [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Completion Message Type column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'CompletionMessageType'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							CompletionMessageType [smallint] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.CompletionMessageType = 0
							WHERE ASRSysWorkflowElements.CompletionMessageType IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Completion Message column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'CompletionMessage'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							CompletionMessage [varchar] (250) NULL'
		EXEC sp_executesql @NVarCommand
	END
	ELSE
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ALTER COLUMN
					CompletionMessage [varchar] (250) NULL '
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new SavedForLater Message Type column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'SavedForLaterMessageType'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							SavedForLaterMessageType [smallint] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.SavedForLaterMessageType = 0
							WHERE ASRSysWorkflowElements.SavedForLaterMessageType IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new SavedForLater Message column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'SavedForLaterMessage'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							SavedForLaterMessage [varchar] (250) NULL'
		EXEC sp_executesql @NVarCommand
	END
	ELSE
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ALTER COLUMN
					SavedForLaterMessage [varchar] (250) NULL '
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new FollowOnForms Message Type column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'FollowOnFormsMessageType'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							FollowOnFormsMessageType [smallint] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.FollowOnFormsMessageType = 0
							WHERE ASRSysWorkflowElements.FollowOnFormsMessageType IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Completion Message column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'FollowOnFormsMessage'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							FollowOnFormsMessage [varchar] (250) NULL'
		EXEC sp_executesql @NVarCommand
	END
	ELSE
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ALTER COLUMN
					FollowOnFormsMessage [varchar] (250) NULL '
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Attachment_Type column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'Attachment_Type'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							Attachment_Type [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.Attachment_Type = -1
							WHERE ASRSysWorkflowElements.Attachment_Type IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Attachment_File column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'Attachment_File'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							Attachment_File [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.Attachment_File = ''''
							WHERE ASRSysWorkflowElements.Attachment_File IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Attachment_WFElementIdentifier column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'Attachment_WFElementIdentifier'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							Attachment_WFElementIdentifier [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.Attachment_WFElementIdentifier = ''''
							WHERE ASRSysWorkflowElements.Attachment_WFElementIdentifier IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Attachment_WFValueIdentifier column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'Attachment_WFValueIdentifier'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							Attachment_WFValueIdentifier [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.Attachment_WFValueIdentifier = ''''
							WHERE ASRSysWorkflowElements.Attachment_WFValueIdentifier IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Attachment_DBColumnID column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'Attachment_DBColumnID'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							Attachment_DBColumnID [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.Attachment_DBColumnID = 0
							WHERE ASRSysWorkflowElements.Attachment_DBColumnID IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Attachment_DBRecord column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'Attachment_DBRecord'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							Attachment_DBRecord [int] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.Attachment_DBRecord = 0
							WHERE ASRSysWorkflowElements.Attachment_DBRecord IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Attachment_DBElement column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'Attachment_DBElement'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							Attachment_DBElement [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.Attachment_DBElement = ''''
							WHERE ASRSysWorkflowElements.Attachment_DBElement IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowElements - Add new Attachment_DBValue column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElements', 'U')
	AND name = 'Attachment_DBValue'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
							Attachment_DBValue [varchar] (200) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElements
							SET ASRSysWorkflowElements.Attachment_DBValue = ''''
							WHERE ASRSysWorkflowElements.Attachment_DBValue IS NULL'
		EXEC sp_executesql @NVarCommand
	END


	/* ASRSysWorkflowElementItems - Add new Password Type column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowElementItems', 'U')
	AND name = 'PasswordType'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
							PasswordType [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysWorkflowElementItems
							SET ASRSysWorkflowElementItems.PasswordType = 0
							WHERE ASRSysWorkflowElementItems.PasswordType IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Drop the FileUpload column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowInstanceValues', 'U')
	AND name = 'FileUpload'

	IF @iRecCount > 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues DROP COLUMN  
							FileUpload '
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new FileUpload_File column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowInstanceValues', 'U')
	AND name = 'FileUpload_File'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
							FileUpload_File [image] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new FileUpload_ContentType column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowInstanceValues', 'U')
	AND name = 'FileUpload_ContentType'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
							FileUpload_ContentType [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysWorkflowInstanceValues - Add new FileUpload_FileName column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysWorkflowInstanceValues', 'U')
	AND name = 'FileUpload_FileName'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
							FileUpload_FileName [varchar] (8000) NULL'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */


/* ------------------------------------------------------------- */
PRINT 'Step 2 of X - Modifying Workflow stored procedures'
/* ------------------------------------------------------------- */

	----------------------------------------------------------------------
	-- spASRWorkflowStoredDataFile
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowStoredDataFile]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowStoredDataFile]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowStoredDataFile]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRWorkflowStoredDataFile]
		(
			@piElementColumnID	integer,
			@piInstanceID		integer,
			@piValueType		integer			OUTPUT,
			@psFileName			varchar(8000)	OUTPUT,
			@psErrorMessage		varchar(8000)	OUTPUT,
			@piOLEType			integer			OUTPUT
		)
		AS
		BEGIN
			DECLARE 
				@iWorkflowID		integer,
				@iElementID			integer,
				@sElementIdentifier	varchar(8000),
				@sItemIdentifier	varchar(8000),
				@iDBColumnID		integer,
				@iDBRecord			integer,
				@sTableName			sysname,
				@sColumnName		sysname,
				@iRequiredTableID	integer,
				@iRequiredRecordID	integer,
				@iRecordID			integer,
				@iBaseTableID		integer,
				@iBaseRecordID		integer,
				@iParent1TableID	int,
				@iParent1RecordID	int,
				@iParent2TableID	int,
				@iParent2RecordID	int,
				@iInitiatorID		integer,
				@iInitParent1TableID	int,
				@iInitParent1RecordID	int,
				@iInitParent2TableID	int,
				@iInitParent2RecordID	int,
				@iElementType		integer, 
				@iTempElementID		integer,
				@sValue				varchar(8000),
				@fValidRecordID		bit,
				@fDeletedValue		bit,
				@iPersonnelTableID	integer,
				@iCount				integer,
				@sSQL				nvarchar(4000),
				@sSQLParam			nvarchar(4000)
		
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
			WHERE WEC.ID = @piElementColumnID
		
			IF @piValueType = 2 -- DB File
			BEGIN
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
					@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
					@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
					@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
		
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
		
				SET @fDeletedValue = 0
		
				SELECT @sTableName = ASRSysTables.tableName, 
					@iRequiredTableID = ASRSysTables.tableID, 
					@sColumnName = ASRSysColumns.columnName,
					@piOLEType = ASRSysColumns.OLEType
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
				END
		
				IF @iDBRecord = 1
				BEGIN
					-- Identified record.
					SELECT @iElementType = ASRSysWorkflowElements.type, 
						@iTempElementID = ASRSysWorkflowElem'


	SET @sSPCode_1 = 'ents.ID
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
						AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sElementIdentifier)))
						
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
							AND Es.identifier = @sElementIdentifier
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
						IF @iDBRecord = 4 -- Trigger record. 
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
						SET @psErrorMessage = ''Record has been deleted or not selected.''
						RETURN
					END
				END
					
				IF @fDeletedValue = 0
				BEGIN
					I'


	SET @sSPCode_2 = 'F (@piOLEType = 0) OR (@piOLEType = 1)
					BEGIN
						SET @sSQL = ''SELECT @psFileName = '' + @sTableName + ''.'' + @sColumnName +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
						SET @sSQLParam = N''@psFileName varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL, @sSQLParam, @psFileName OUTPUT
					END
					ELSE
					BEGIN
						SET @sSQL = ''SELECT '' + @sTableName + ''.'' + @sColumnName + '' AS [file]'' +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
						EXEC sp_executesql @sSQL
					END
				END
			END
			
			IF @piValueType = 1 -- WF File
			BEGIN
				SELECT @iElementID = isnull(ID, 0)
				FROM ASRSysWorkflowElements
				WHERE workflowID = @iWorkflowID
					AND upper(ltrim(rtrim(isnull(identifier, '''')))) = @sElementIdentifier
		
				SELECT @psFileName = fileUpload_fileName
				FROM ASRSysWorkflowInstanceValues
				WHERE instanceID = @piInstanceID
					AND elementID = @iElementID
					AND upper(ltrim(rtrim(isnull(identifier, '''')))) = @sItemIdentifier
		
				SELECT fileUpload_file AS [file]
				FROM ASRSysWorkflowInstanceValues
				WHERE instanceID = @piInstanceID
					AND elementID = @iElementID
					AND upper(ltrim(rtrim(isnull(identifier, '''')))) = @sItemIdentifier
			END
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

	----------------------------------------------------------------------
	-- spASRWorkflowValidRecord
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowValidRecord]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowValidRecord]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowValidRecord]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRWorkflowValidRecord
					@piInstanceID				integer,
					@piRecordType				integer,
					@piRecordID					integer,
					@sElementIdentifier			varchar(8000),
					@sElementItemIdentifier		varchar(8000),
					@pfValid					bit				OUTPUT
				AS
				BEGIN
					DECLARE
						@sSQL					nvarchar(4000),
						@iTableID				integer,
						@sTableName				nvarchar(4000),
						@iWorkflowID			integer,
						@sParam					nvarchar(500),
						@iRecCount				integer,
						@iElementType			integer
				
					SET @pfValid = 0
				
					SELECT @iWorkflowID = WF.ID,
						@iTableID = 
							CASE
								WHEN @piRecordType = 4 THEN isnull(WF.baseTable, 0)
								ELSE 0
							END
					FROM ASRSysWorkflows WF
					INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
						AND WFI.ID = @piInstanceID
				
					IF @piRecordType = 0
					BEGIN
						-- Initiator''s record
						SELECT @iTableID = convert(integer, ISNULL(parameterValue, ''0''))
						FROM ASRSysModuleSetup
						WHERE moduleKey = ''MODULE_PERSONNEL''
							AND parameterKey = ''Param_TablePersonnel''
		
						IF @iTableID = 0
						BEGIN
							SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
							FROM ASRSysModuleSetup
							WHERE moduleKey = ''MODULE_WORKFLOW''
							AND parameterKey = ''Param_TablePersonnel''
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
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sElementIdentifier)))
				
						IF @iElementType = 2
						BEGIN
							 -- WebForm
							SELECT @iTableID = WFEI.tableID
							FROM ASRSysWorkflowElementItems WFEI
							INNER JOIN ASRSysWorkflowElements WFE ON WFEI.elementID = WFE.ID
								AND WFE.identifier = @sElementIdentifier
								AND WFE.workflowID = @iWorkflowID
							WHERE WFEI.identifier = @sElementItemIdentifier
						END
					END
				
					exec spASRWorkflowValidTableRecord
						@iTableID,
						@piRecordID,
						@pfValid	OUTPUT
				END'

	EXECUTE (@sSPCode_0)

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
										SELECT @iBaseTableID = convert(integer, ISNULL(parameterValue, ''0''))
										FROM ASRSysModuleSetup
										WHERE moduleKey = ''MODULE_PERSONNEL''
											AND parameterKey = ''Param_TablePersonnel''

										IF @iBaseTableID = 0
										BEGIN
											SELECT @iBaseTableID = convert(integer, i'


	SET @sSPCode_3 = 'snull(parameterValue, 0))
											FROM ASRSysModuleSetup
											WHERE moduleKey = ''MODULE_WORKFLOW''
											AND parameterKey = ''Param_TablePersonnel''
										END
									END
								END
		
								IF @iEmailRecord = 1
								BEGIN
									SELECT @iPrevElementType = ASRSysWorkflowElements.type,
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

												IF len(@sTempTo) > 0 '


	SET @sSPCode_4 = 'SET @fValidRecordID = 1
											END
										END
									END

									IF (@fValidRecordID = 0) AND (@iEmailLoop = 0)
									BEGIN
										-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
										EXEC [dbo].[spASRWorkflowActionFailed] 
											@piInstanceID, 
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
					ASRSysWorkflowInstanceSteps.h'


	SET @sSPCode_5 = 'ypertextLinkedSteps = CASE
						WHEN @iElementType = 3 THEN @sHypertextLinkedSteps
						ELSE ASRSysWorkflowInstanceSteps.hypertextLinkedSteps
					END,
					ASRSysWorkflowInstanceSteps.message = CASE
						WHEN @iElementType = 3 THEN LEFT(@sMessage, 8000)
						WHEN @iElementType = 5 THEN LEFT(@sMessage, 8000)
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

						-- If the submitted element is a web form, then any succeeding webforms are actioned f'


	SET @sSPCode_6 = 'or the same user.
						IF @iElementType = 2 -- WebForm
						BEGIN
							SELECT @sUserName = isnull(WIS.userName, ''''),
								@sUserEmail = isnull(WIS.userEmail, '''')
							FROM ASRSysWorkflowInstanceSteps WIS
							WHERE WIS.instanceID = @piInstanceID
								AND WIS.elementID = @piElementID

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
					'


	SET @sSPCode_7 = '					OR ASRSysWorkflowInstanceSteps.status = 2
										OR ASRSysWorkflowInstanceSteps.status = 6
										OR ASRSysWorkflowInstanceSteps.status = 8
										OR ASRSysWorkflowInstanceSteps.status = 3))
							
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
					UP'


	SET @sSPCode_8 = 'DATE ASRSysWorkflowInstanceSteps
					SET ASRSysWorkflowInstanceSteps.status = 0 -- 0 = On hold
					WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
						AND (ASRSysWorkflowInstanceSteps.status = 1 -- 1 = Pending Engine Action
							OR ASRSysWorkflowInstanceSteps.status = 2) -- 2 = Pending User Action
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
			
				IF @iPersonnelTableID = 0
				BEGIN
					SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_WORKFLOW''
					AND parameterKey = ''Param_TablePersonnel''
				END
							
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
					ASRSysOrder'


	SET @sSPCode_1 = 'Items.sequence
			
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
							FROM ASRSysWorkflowInstanc'


	SET @sSPCode_2 = 'eValues IV
							INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
								AND IV.identifier = Es.identifier
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sRecSelWebFormIdentifier
							WHERE IV.instanceID = @piInstanceID
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
									@iPersonnelTableID	integer,
									@iSQLVersion	integer
											
								SET @pfOK = 1
								SET @psMessage = ''''
								SET @psMessage_HypertextLinks = ''''
								SET @psHypertextLinkedSteps = ''''
								SELECT @iSQLVersion = dbo.udfASRSQLVersion()
							
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
									@iInitParent1TableID'


	SET @sSPCode_1 = ' = ASRSysWorkflowInstances.parent1TableID,
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
									EI.wfFormIdentifier,
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
									'


	SET @sSPCode_2 = '			SELECT @sTemp = ISNULL(IV.value, ''0''),
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
							
				'


	SET @sSPCode_3 = '						/* Format dates */
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
				'


	SET @sSPCode_4 = '							@dtResult OUTPUT,
											@fltResult OUTPUT, 
											0
		
										SET @psMessage = @psMessage +
											@sResult
									END
							
									FETCH NEXT FROM itemCursor INTO @sCaption, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier, @sRecSelWebFormIdentifier, @sRecSelIdentifier, @iCalcID
								END
								CLOSE itemCursor
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
'


	SET @sSPCode_5 = '
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
										+ CASE
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

				IF @iPersonnelTableID = 0
				BEGIN
					SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_WORKFLOW''
					AND parameterKey = ''Param_TablePersonnel''
				END
			
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
				INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRS'


	SET @sSPCode_1 = 'ysWorkflows.ID
				WHERE ASRSysWorkflowElements.ID = @piElementID
			
				SELECT @psTableName = tableName
				FROM ASRSysTables
				WHERE tableID = @piDataTableID
			
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
				'


	SET @sSPCode_2 = '		-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
						EXEC [dbo].[spASRWorkflowActionFailed]
							@piInstanceID, 
							@piElementID, 
							''Stored Data primary record has been deleted or not selected.''

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

						'


	SET @sSPCode_3 = 'IF @piDataTableID = @iSecondaryDataRecordTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iSecondaryDataRecordTableID)
						END
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
							-- Identified record
							SELECT @iElementType = ASRSysWorkflowElements.type, 
								@iTempElementID ='


	SET @sSPCode_4 = ' ASRSysWorkflowElements.ID
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

											SET @'


	SET @sSPCode_5 = 'fValidRecordID = 1
											SET @fDeletedValue = 1
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
									''Stored Data column database value record has been deleted or not selected.''

								SET @psSQL = ''''
								RETURN
							END
						END

						IF (@iDataType <> -3)
							AND (@iDataType <> -4)
						BEGIN
							IF @fDeletedValue = 0
							BEGIN
								SET @sSQL = @sSQL + convert(nvarchar(4000), @iRecordID)
								SET @sParam = N''@sDBValue varchar(8000) OUTPUT''
								EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT
							END

							UPDATE @dbValues
							SET value = @sDBValue
							WHERE ID = @iID
						END
						
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
						AND ((SC.dataType <> -3) AND (SC.dataType <> -4))
			
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
			'


	SET @sSPCode_6 = '				/* INSERT. */
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
							AND columnID = @iColumnID

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
			
				IF @piDataAc'


	SET @sSPCode_7 = 'tion = 2
				BEGIN
					/* DELETE. */
					SET @psSQL = ''DELETE FROM '' + @psTableName
						+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(8000), @piRecordID)
				END	

				IF (@piDataAction = 0) -- Insert
				BEGIN
					SET @iParent1TableID = isnull(@iDataRecordTableID, 0)
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
									WHEN '


	SET @sSPCode_8 = '@iDataType = 12 THEN ''''
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
				@piHeight	integer	OUTPUT,
				@piCompletionMessageType	integer	OUTPUT,
				@psCompletionMessage		varchar(200)	OUTPUT,
				@piSavedForLaterMessageType	integer	OUTPUT,
				@psSavedForLaterMessage		varchar(200)	OUTPUT,
				@piFollowOnFormsMessageType	integer	OUTPUT,
				@psFollowOnFormsMessage		varchar(200)	OUTPUT
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
					@iPersonnelTableID	integer,
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
					@piCompletio'


	SET @sSPCode_1 = 'nMessageType = CompletionMessageType,
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
							-- Identified record'


	SET @sSPCode_2 = '.
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
					'


	SET @sSPCode_3 = '							AND IV.elementID = @iTempElementID

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
									'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
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
								EXEC [dbo].[spASRSysWorkfl'


	SET @sSPCode_4 = 'owCalculation]
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
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4)

	----------------------------------------------------------------------
	-- spASRInstantiateWorkflow
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRInstantiateWorkflow]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRInstantiateWorkflow]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRInstantiateWorkflow]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASRInstantiateWorkflow]
		(
			@piWorkflowID	integer,			
			@piInstanceID	integer			OUTPUT,
			@psFormElements	varchar(8000)	OUTPUT,
			@psMessage		varchar(8000)	OUTPUT
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
				@fUsesInitiator	bit, 
				@iTemp int,
				@iStartElementID int,
				@iTableID				integer,
				@iParent1TableID		integer,
				@iParent1RecordID		integer,
				@iParent2TableID		integer,
				@iParent2RecordID		integer,
				@sForms	varchar(8000),
				@iCount	integer,
				@iSQLVersion			integer,
				@fExternallyInitiated	bit,
				@fEnabled	bit,
				@iElementType	integer,
				@fStoredDataOK			bit, 
				@sStoredDataMsg			varchar(8000), 
				@sStoredDataSQL			varchar(8000), 
				@iStoredDataTableID		integer,
				@sStoredDataTableName	varchar(8000),
				@iStoredDataAction		integer, 
				@iStoredDataRecordID	integer,
				@sStoredDataRecordDesc	varchar(8000),
				@sSPName	varchar(8000),
				@iNewRecordID	integer,
				@sEvalRecDesc	varchar(8000),
				@iResult	integer,
				@iFailureFlows	integer,
				@fSaveForLater bit	

			SELECT @iSQLVersion = convert(float,substring(@@version,charindex(''-'',@@version)+2,2))

			DECLARE @succeedingElements table(elementID int)
		
			SET @iInitiatorID = 0
			SET @psFormElements = ''''
			SET @psMessage = ''''
			SET @iParent1TableID = 0
			SET @iParent1RecordID = 0
			SET @iParent2TableID = 0
			SET @iParent2RecordID = 0
		
			SELECT @fExternallyInitiated = CASE
					WHEN initiationType = 2 THEN 1
					ELSE 0
				END,
				@fEnabled = enabled
			FROM ASRSysWorkflows
			WHERE ID = @piWorkflowID

			IF @fExternallyInitiated = 1
			BEGIN
				IF @fEnabled = 0
				BEGIN
					/* Workflow is disabled. */
					SET @psMessage = ''This link is currently disabled.''
					RETURN
				END

				SET @sActualLoginName = ''<External>''
			END
			ELSE
			BEGIN
				SET @sActualLoginName = SUSER_SNAME()
				
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
					EXEC [dbo].[spASRWorkflowUsesInitiator]
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
				ELSE
				BEGIN
					SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
					FROM ASRSysModuleSetup
					WHERE moduleKey = ''MODULE_PERSONNEL''
					AND parameterKey = ''Param_TablePersonnel''

					IF @iTableID = 0 
					BEGIN
						SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
						FROM ASRSysModuleSetup
						WHERE moduleKey = ''MODULE_WORKFLOW''
						AND parameterKey = ''Param_TablePersonnel''
					END

					exec [dbo].[spASRGetParentDetails]
						@iTableID,
						@iInitiatorID,
						@iParent1TableID	OUTPUT,
						@iParent1RecordID	OUTPUT,
						@iParent2TableID	OUTPUT,
						@iParent2RecordID	OUTPUT
				END
			END
		
			/* Create the Workflow Instance record, and remember the ID. */
			INSERT INTO ASRSysWorkflowInstances (workflowID, 
				initiatorID, 
				status, 
				userName, 
				pare'


	SET @sSPCode_1 = 'nt1TableID,
				parent1RecordID,
				parent2TableID,
				parent2RecordID)
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID)
						
			SELECT @piInstanceID = MAX(id)
			FROM ASRSysWorkflowInstances
		
			/* Create the Workflow Instance Steps records. 
			Set the first steps'' status to be 1 (pending Workflow Engine action). 
			Set all subsequent steps'' status to be 0 (on hold). */

			SELECT @iStartElementID = ASRSysWorkflowElements.ID
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.type = 0 -- Start element
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID

			INSERT INTO @succeedingElements 
			SELECT id 
			FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0)
		
			INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
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
			WHERE ASRSysWorkflowElements.workflowid = @piWorkflowID
		
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
				AND ASRSysWorkflowElements.type = 5
						
			SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
						OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		
			WHILE @iCount > 0 
			BEGIN
				DECLARE immediateSubmitCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ASRSysWorkflowInstanceSteps.elementID, 
					ASRSysWorkflowElements.type
				FROM ASRSysWorkflowInstanceSteps
				INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowInstanceSteps.status = 1
					AND (ASRSysWorkflowElements.type = 4 
		'


	SET @sSPCode_2 = '				OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
						OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
					AND ASRSysWorkflowElements.workflowID = @piWorkflowID
					AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		

				OPEN immediateSubmitCursor
				FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, 
					@iElementType 
				WHILE (@@fetch_status = 0) 
				BEGIN
					IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
					BEGIN
						SET @fStoredDataOK = 1
						SET @sStoredDataMsg = ''''
						SET @sStoredDataRecordDesc = ''''

						EXEC [spASRGetStoredDataActionDetails]
							@piInstanceID,
							@iElementID,
							@sStoredDataSQL			OUTPUT, 
							@iStoredDataTableID		OUTPUT,
							@sStoredDataTableName	OUTPUT,
							@iStoredDataAction		OUTPUT, 
							@iStoredDataRecordID	OUTPUT

						IF @iStoredDataAction = 0 -- Insert
						BEGIN
							--SET @sSPName  = ''sp_ASRInsertNewRecord_'' + convert(varchar(8000), @iStoredDataTableID)						 
							SET @sSPName  = ''spASRWorkflowInsertNewRecord''

							BEGIN TRY
								EXEC @sSPName
									@iNewRecordID  OUTPUT, 
									@iStoredDataTableID,
									@sStoredDataSQL

								SET @iStoredDataRecordID = @iNewRecordID
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0
								SET @sStoredDataMsg = ERROR_MESSAGE()
							END CATCH
						END
						ELSE IF @iStoredDataAction = 1 -- Update
						BEGIN
							--SET @sSPName  = ''sp_ASRUpdateRecord_'' + convert(varchar(8000), @iStoredDataTableID)
							SET @sSPName  = ''spASRWorkflowUpdateRecord''

							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataSQL,
									@sStoredDataTableName,
									@iStoredDataRecordID
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0
								SET @sStoredDataMsg = ERROR_MESSAGE()
							END CATCH
						END
						ELSE IF @iStoredDataAction = 2 -- Delete
						BEGIN
							EXEC spASRRecordDescription
								@iStoredDataTableID,
								@iStoredDataRecordID,
								@sStoredDataRecordDesc OUTPUT

							--SET @sSPName  = ''sp_ASRDeleteRecord_'' + convert(varchar(8000), @iStoredDataTableID)
							SET @sSPName  = ''spASRWorkflowDeleteRecord''

							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
									@sStoredDataTableName,
									@iStoredDataRecordID
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0
								SET @sStoredDataMsg = ERROR_MESSAGE()
							END CATCH
						END
						ELSE
						BEGIN
							SET @fStoredDataOK = 0
							SET @sStoredDataMsg = ''Unrecognised data action.''
						END

						IF (@fStoredDataOK = 1)
							AND ((@iStoredDataAction = 0)
								OR (@iStoredDataAction = 1))
						BEGIN

							exec [dbo].[spASRStoredDataFileActions]
								@piInstanceID,
								@iElementID,
								@iStoredDataRecordID
						END

						IF @fStoredDataOK = 1
						BEGIN
							SET @sStoredDataMsg = ''Successfully '' +
								CASE
									WHEN @iStoredDataAction = 0 THEN ''inserted''
									WHEN @iStoredDataAction = 1 THEN ''updated''
									ELSE ''deleted''
								END + '' record''

							IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
							BEGIN
								IF @iStoredDataRecordID > 0 
								BEGIN	
									EXEC [dbo].[spASRRecordDescription] 
										@iStoredDataTableID,
										@iStoredDataRecordID,
										@sEvalRecDesc OUTPUT
									IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc
								END
							END

							IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')''

							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = convert(varchar(8000), @iStoredDataRec'


	SET @sSPCode_3 = 'ordID), 
								ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @iElementID
								AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
								AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0

							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 3,
								ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.message = LEFT(@sStoredDataMsg, 8000)
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID = @iElementID

							-- Get this immediate element''s succeeding elements
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 1
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
									FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC)
						END
						ELSE
						BEGIN
							-- Check if the failed element has an outbound flow for failures.
							SELECT @iFailureFlows = COUNT(*)
							FROM ASRSysWorkflowElements Es
							INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
								AND Ls.startOutboundFlowCode = 1
							WHERE Es.ID = @iElementID
								AND Es.type = 5 -- 5 = StoredData

							IF @iFailureFlows = 0
							BEGIN
								UPDATE ASRSysWorkflowInstanceSteps
								SET status = 4,	-- 4 = failed
									message = @sStoredDataMsg,
									failedCount = isnull(failedCount, 0) + 1,
									completionCount = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID

								UPDATE ASRSysWorkflowInstances
								SET status = 2	-- 2 = error
								WHERE ID = @piInstanceID

								SET @psMessage = @sStoredDataMsg
								RETURN
							END
							ELSE
							BEGIN
								UPDATE ASRSysWorkflowInstanceSteps
								SET status = 8,	-- 8 = failed action
									message = @sStoredDataMsg,
									failedCount = isnull(failedCount, 0) + 1,
									completionCount = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID

								UPDATE ASRSysWorkflowInstanceSteps
								SET ASRSysWorkflowInstanceSteps.status = 1
								WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
									AND ASRSysWorkflowInstanceSteps.elementID IN (SELECT SUCC.id
										FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, 0) SUCC)
							END
						END
					END
					ELSE
					BEGIN
						EXEC [dbo].[spASRSubmitWorkflowStep] 
							@piInstanceID, 
							@iElementID, 
							'''', 
							'''', 
							@sForms OUTPUT, 
							@fSaveForLater OUTPUT
					END

					FETCH NEXT FROM immediateSubmitCursor INTO @iElementID, 
						@iElementType
				END
				CLOSE immediateSubmitCursor
				DEALLOCATE immediateSubmitCursor

				SELECT @iCount = COUNT(ASRSysWorkflowInstanceSteps.elementID)
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
						AND (ASRSysWorkflowElements.type = 4 
							OR (@iSQLVersion >= 9 AND ASRSysWorkflowElements.type = 5) 
							OR ASRSysWorkflowElements.type = 7) -- 4=Decision, 5=StoredData, 7=Or
						AND ASRSysWorkflowElements.workflowID = @piWorkflowID
						AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		
			END						

			/* Return a list of the workflow form elements that may need to be displayed to the initiator straight away */
			DECLARE formsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ASRSysWorkflowInstanceSteps.ID,
				ASRSysWorkflowInstanceSteps.elementID
			FROM ASRSysWorkflowInstanceSteps'


	SET @sSPCode_4 = '
			INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
			WHERE (ASRSysWorkflowInstanceSteps.status = 1 OR ASRSysWorkflowInstanceSteps.status = 2)
				AND ASRSysWorkflowElements.type = 2
				AND ASRSysWorkflowElements.workflowID = @piWorkflowID
				AND ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID		
		
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

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4)


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
			-- Action any immediate elements (Or, Decision and StoredData elements) and return the IDs of the workflow elements that 
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
				@iSQLVersion			integer,
				@fStoredDataOK			bit, 
				@sStoredDataMsg			varchar(8000), 
				@sStoredDataSQL			varchar(8000), 
				@iStoredDataTableID		integer,
				@sStoredDataTableName	varchar(8000),
				@iStoredDataAction		integer, 
				@iStoredDataRecordID	integer,
				@sStoredDataRecordDesc	varchar(8000),
				@sStoredDataWebForms	varchar(8000),
				@sStoredDataSaveForLater bit,
				@sSPName	varchar(8000),
				@iNewRecordID	integer,
				@sEvalRecDesc	varchar(8000),
				@iResult	integer,
				@iFailureFlows	integer,
				@fIsDelegate		bit
		
			SELECT @iSQLVersion = convert(float,substring(@@version,charindex(''-'',@@version)+2,2))
							
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
			WHERE (elementType = 4 OR (@iSQLVersion >= 9 AND elementType = 5) OR elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
				AND processed = 0
		
			WHILE @iCount > 0
			BEGIN
				UPDATE @elements
				SET processed = 1
				WHERE processed = 0
		
				-- Action any succeeding immediate elements (Decision, Or and StoredData elements)
				DECLARE immediateCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT E.elementID,
					E.elementType,
					E.trueFlowType, 
					E.trueFlowExprID
				FROM @elements E
				WHERE (E.elementType = 4 OR (@iSQLVersion >= 9 AND E.elementType = 5) OR E.elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
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
						ASRSysWorkflowInstanceSteps.activationDateTime = getdate(), 
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
					'


	SET @sSPCode_1 = '			@piInstanceID,
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
								INNER JOIN ASRSysWorkflowInstanceSteps WIS ON PE.ID = WIS.elementID
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
					ELSE IF (@iElementType = 5) AND (@iSQLVersion >= 9) -- StoredData
					BEGIN
						SET @fStoredDataOK = 1
						SET @sStoredDataMsg = ''''
						SET @sStoredDataRecordDesc = ''''
		
						EXEC [spASRGetStoredDataActionDetails]
							@piInstanceID,
							@iElementID,
							@sStoredDataSQL			OUTPUT, 
							@iStoredDataTableID		OUTPUT,
							@sStoredDataTableName	OUTPUT,
							@iStoredDataAction		OUTPUT, 
							@iStoredDataRecordID	OUTPUT
		
						IF @iStoredDataAction = 0 -- Insert
						BEGIN
							SET @sSPName  = ''sp_ASRInsertNewRecord_'' + convert(varchar(8000), @iStoredDataTableID)
		
							BEGIN TRY
								EXEC @sSPName
									@iNewRecordID  OUTPUT, 
									@sStoredDataSQL
		
								SET @iStoredDataRecordID = @iNewRecordID
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0
								SET @sStoredDataMsg = ERROR_MESSAGE()
							END CATCH
						END
						ELSE IF @iStoredDataAction = 1 -- Update
						BEGIN
							SET @sSPName  = ''sp_ASRUpdateRecord_'' + convert(varchar(8000), @iStoredDataTableID)
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@sStoredDataSQL,
									@sStoredDataTableName,
									@iStoredDataRecordID,
									null
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0
								SET @sStoredDataMsg = ERROR_MESSAGE()
							END CATCH
						END
						ELSE IF @iStoredDataAction = 2 -- Delete
						BEGIN
							EXEC spASRRecordDescription
								@iStoredDataTableID,
								@iStoredDataRecordID,
								@sStoredDataRecordDesc OUTPUT
		
							SET @sSPName  = ''sp_ASRDeleteRecord_'' + convert(varchar(8000), @iStoredDataTableID)
		
							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@sStoredDataTableName,
									@iStoredDataRecordID
							END TRY
							BEGIN CATCH
								SET @fStoredDataOK = 0
								SET @sStoredDataMsg = ERROR_MESSAGE()
							END CATCH
						END
						ELSE
						BEGIN
							SET @fStoredDataOK = 0
							SET @sStoredDataMsg = ''Unrecognised data action.''
				'


	SET @sSPCode_2 = '		END
		
						IF (@fStoredDataOK = 1)
							AND ((@iStoredDataAction = 0)
								OR (@iStoredDataAction = 1))
						BEGIN
		
							exec [dbo].[spASRStoredDataFileActions]
								@piInstanceID,
								@iElementID,
								@iStoredDataRecordID
						END
		
						IF @fStoredDataOK = 1
						BEGIN
							SET @sStoredDataMsg = ''Successfully '' +
								CASE
									WHEN @iStoredDataAction = 0 THEN ''inserted''
									WHEN @iStoredDataAction = 1 THEN ''updated''
									ELSE ''deleted''
								END + '' record''
		
							IF (@iStoredDataAction = 0) OR (@iStoredDataAction = 1) -- Inserted or Updated
							BEGIN
								IF @iStoredDataRecordID > 0 
								BEGIN	
									EXEC [dbo].[spASRRecordDescription] 
										@iStoredDataTableID,
										@iStoredDataRecordID,
										@sEvalRecDesc OUTPUT
									IF (NOT @sEvalRecDesc IS null) AND (LEN(@sEvalRecDesc) > 0) SET @sStoredDataRecordDesc = @sEvalRecDesc
								END
							END
		
							IF len(@sStoredDataRecordDesc) > 0 SET @sStoredDataMsg = @sStoredDataMsg + '' ('' + @sStoredDataRecordDesc + '')''
		
							UPDATE ASRSysWorkflowInstanceValues
							SET ASRSysWorkflowInstanceValues.value = convert(varchar(8000), @iStoredDataRecordID), 
								ASRSysWorkflowInstanceValues.valueDescription = @sStoredDataRecordDesc
							WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceValues.elementID = @iElementID
								AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
								AND isnull(ASRSysWorkflowInstanceValues.emailID, 0) = 0
		
							UPDATE ASRSysWorkflowInstanceSteps
							SET ASRSysWorkflowInstanceSteps.status = 3,
								ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
								ASRSysWorkflowInstanceSteps.message = LEFT(@sStoredDataMsg, 8000)
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID = @iElementID
						END
						ELSE
						BEGIN
							-- Check if the failed element has an outbound flow for failures.
							SELECT @iFailureFlows = COUNT(*)
							FROM ASRSysWorkflowElements Es
							INNER JOIN ASRSysWorkflowLinks Ls ON Es.ID = Ls.startElementID
								AND Ls.startOutboundFlowCode = 1
							WHERE Es.ID = @iElementID
								AND Es.type = 5 -- 5 = StoredData
		
							IF @iFailureFlows = 0
							BEGIN
								UPDATE ASRSysWorkflowInstanceSteps
								SET status = 4,	-- 4 = failed
									message = @sStoredDataMsg,
									failedCount = isnull(failedCount, 0) + 1,
									completionCount = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID
		
								UPDATE ASRSysWorkflowInstances
								SET status = 2	-- 2 = error
								WHERE ID = @piInstanceID
							END
							ELSE
							BEGIN
								UPDATE ASRSysWorkflowInstanceSteps
								SET status = 8,	-- 8 = failed action
									message = @sStoredDataMsg,
									failedCount = isnull(failedCount, 0) + 1,
									completionCount = isnull(completionCount, 0) - 1
								WHERE instanceID = @piInstanceID
									AND elementID = @iElementID
		
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
								WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements)
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
						SELECT SUCC.id'


	SET @sSPCode_3 = ',
							E.type,
							0,
							isnull(E.trueFlowType, 0),
							isnull(E.trueFlowExprID, 0)
						FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iElementID, @iFlowCode) SUCC
						INNER JOIN ASRSysWorkflowElements E ON SUCC.ID = E.ID
						WHERE SUCC.ID NOT IN (SELECT elementID FROM @elements)
					END
		
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
				WHERE (elementType = 4 OR (@iSQLVersion >= 9 AND elementType = 5) OR elementType = 7) -- 4=Decision, 5=StoredData, 7=Or
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
						AND (E.elementType <> 5 OR @iSQLVersion <= 8) -- 5 = StoredData'


	SET @sSPCode_4 = '
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
					AND (E.elementType <> 5 OR @iSQLVersion <= 8) -- 5 = StoredData
		
			OPEN @succeedingElements
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3
		+ @sSPCode_4)

	----------------------------------------------------------------------
	-- spASRGetWorkflowFileUploadDetails
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowFileUploadDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowFileUploadDetails]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowFileUploadDetails]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetWorkflowFileUploadDetails]
		(
			@piElementItemID	integer,
			@piInstanceID		integer,
			@piSize				integer			OUTPUT,
			@psFileName			varchar(8000)	OUTPUT
		)
		AS
		BEGIN
			DECLARE
				@iElementID	integer,
				@sIdentifier	varchar(8000) 
		
			SELECT 			
				@piSize = ISNULL(ASRSysWorkflowElementItems.InputSize, 0),
				@iElementID = elementID,
				@sIdentifier = identifier
			FROM ASRSysWorkflowElementItems
			WHERE ASRSysWorkflowElementItems.ID = @piElementItemID
		
			SELECT @psFileName = [FileUpload_Filename]
			FROM ASRSysWorkflowInstanceValues
			WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceValues.elementID = @iElementID
				AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier
		
			SELECT ASRSysWorkflowElementItemValues.value
			FROM ASRSysWorkflowElementItemValues
			WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRInstantiateTriggeredWorkflows
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRInstantiateTriggeredWorkflows]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRInstantiateTriggeredWorkflows
		AS
		BEGIN
			DECLARE
				@iQueueID			integer,
				@iWorkflowID		integer,
				@iRecordID			integer,
				@iInstanceID		integer,
				@iStartElementID	integer,
				@iTemp				integer,
				@iParent1TableID	integer,
				@iParent1RecordID	integer,
				@iParent2TableID	integer,
				@iParent2RecordID	integer

			DECLARE @succeedingElements table(elementID int)
		
			DECLARE triggeredWFCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT Q.queueID,
				Q.recordID,
				TL.workflowID,
				Q.parent1TableID,
				Q.parent1RecordID,
				Q.parent2TableID,
				Q.parent2RecordID
			FROM ASRSysWorkflowQueue Q
			INNER JOIN ASRSysWorkflowTriggeredLinks TL ON Q.linkID = TL.linkID
			INNER JOIN ASRSysWorkflows WF ON TL.workflowID = WF.ID
				AND WF.enabled = 1
			WHERE Q.dateInitiated IS null
				AND datediff(dd,DateDue,getdate()) >= 0
		
			OPEN triggeredWFCursor
			FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID
			WHILE (@@fetch_status = 0) 
			BEGIN
				UPDATE ASRSysWorkflowQueue
				SET dateInitiated = getDate()
				WHERE queueID = @iQueueID
				
				-- Create the Workflow Instance record, and remember the ID. */
				INSERT INTO ASRSysWorkflowInstances (workflowID, 
					initiatorID, 
					status, 
					userName, 
					parent1TableID,
					parent1RecordID,
					parent2TableID,
					parent2RecordID)
				VALUES (@iWorkflowID, 
					@iRecordID, 
					0, 
					''<Triggered>'',
					@iParent1TableID,
					@iParent1RecordID,
					@iParent2TableID,
					@iParent2RecordID)
								
				SELECT @iInstanceID = MAX(id)
				FROM ASRSysWorkflowInstances
				
				UPDATE ASRSysWorkflowQueue
				SET instanceID = @iInstanceID
				WHERE queueID = @iQueueID	

				-- Create the Workflow Instance Steps records. 
				-- Set the first steps'' status to be 1 (pending Workflow Engine action). 
				-- Set all subsequent steps'' status to be 0 (on hold). */
				SELECT @iStartElementID = ASRSysWorkflowElements.ID
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.type = 0 -- Start element
					AND ASRSysWorkflowElements.workflowID = @iWorkflowID
		
				INSERT INTO @succeedingElements 
				SELECT id 
				FROM [dbo].[udfASRGetSucceedingWorkflowElements](@iStartElementID, 0)
		
				INSERT INTO ASRSysWorkflowInstanceSteps (instanceID, elementID, status, activationDateTime, completionDateTime, completionCount, failedCount, timeoutCount)
				SELECT 
					@iInstanceID, 
					ASRSysWorkflowElements.ID, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN 3
						WHEN ASRSysWorkflowElements.ID IN (SELECT elementID
						FROM @succeedingElements) THEN 1
						ELSE 0
					END, 
					CASE
						WHEN ASRSysWorkflowElements.type = 0 THEN getdate()
						WHEN ASRSysWorkflowElements.ID IN (SELECT elementID
						FROM @succeedingElements) THEN getdate()
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
				WHERE ASRSysWorkflowElements.workflowid = @iWorkflowID
				
				-- Create the Workflow Instance Value records. 
				INSERT INTO ASRSysWorkflowInstanceValues (instanceID, elementID, identifier)
				SELECT @iInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElementItems.identifier
				FROM ASRSysWorkflowElementItems 
				INNER JOIN ASRSysWorkflowElements on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND ASRSysWorkflowElements.type = 2
					AND (ASRSysWorkflowElementItems.itemType = 3 
						OR ASRSysWorkflowElementItems.itemType = 5
						OR ASRSysWorkflowElementItems.itemType = 6
						OR ASRSysWorkflowElement'


	SET @sSPCode_1 = 'Items.itemType = 7
						OR ASRSysWorkflowElementItems.itemType = 11
						OR ASRSysWorkflowElementItems.itemType = 13
						OR ASRSysWorkflowElementItems.itemType = 14
						OR ASRSysWorkflowElementItems.itemType = 15
						OR ASRSysWorkflowElementItems.itemType = 17
						OR ASRSysWorkflowElementItems.itemType = 0)
				UNION
				SELECT  @iInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElements.identifier
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND ASRSysWorkflowElements.type = 5						
				
				FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID
			END
			CLOSE triggeredWFCursor
			DEALLOCATE triggeredWFCursor
		END
		'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

	----------------------------------------------------------------------
	-- spASRWorkflowFileUpload
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowFileUpload]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowFileUpload]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowFileUpload]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRWorkflowFileUpload]
		(
			@piElementItemID	integer,
			@piInstanceID		integer,
			@pimgFile			image,
			@psContentType		varchar(8000),
			@psFileName			varchar(8000),
			@pfClear			bit
		)
		AS
		BEGIN
			DECLARE
				@iElementID	integer,
				@sIdentifier	varchar(8000) 
		
			SELECT
				@iElementID = elementID,
				@sIdentifier = identifier
			FROM ASRSysWorkflowElementItems
			WHERE id = @piElementItemID
		
			UPDATE ASRSysWorkflowInstanceValues 
			SET [FileUpload_File] = 
					CASE 
						WHEN @pfClear = 1 THEN null
						ELSE @pimgFile
					END, 
				[FileUpload_ContentType] = 
					CASE 
						WHEN @pfClear = 1 THEN null
						ELSE @psContentType
					END, 
				[FileUpload_Filename] = 
					CASE 
						WHEN @pfCLear = 1 THEN null
						ELSE @psFileName
					END
			WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceValues.elementID = @iElementID
				AND ASRSysWorkflowInstanceValues.identifier = @sIdentifier
		
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRWorkflowFileDownload
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowFileDownload]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowFileDownload]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowFileDownload]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRWorkflowFileDownload]
		(
			@piElementItemID	integer,
			@piInstanceID		integer,
			@piItemType			integer			OUTPUT,
			@psFileName			varchar(8000)	OUTPUT,
			@psContentType		varchar(8000)	OUTPUT,
			@psErrorMessage		varchar(8000)	OUTPUT,
			@piOLEType			integer			OUTPUT,
			@piDBColumnType		integer			OUTPUT
		)
		AS
		BEGIN
			DECLARE 
				@iWorkflowID		integer,
				@iElementID			integer,
				@sElementIdentifier	varchar(8000),
				@sItemIdentifier	varchar(8000),
				@iDBColumnID		integer,
				@iDBRecord			integer,
				@sTableName			sysname,
				@sColumnName		sysname,
				@iRequiredTableID	integer,
				@iRequiredRecordID	integer,
				@iRecordID			integer,
				@iBaseTableID		integer,
				@iBaseRecordID		integer,
				@iParent1TableID	int,
				@iParent1RecordID	int,
				@iParent2TableID	int,
				@iParent2RecordID	int,
				@iInitiatorID		integer,
				@iInitParent1TableID	int,
				@iInitParent1RecordID	int,
				@iInitParent2TableID	int,
				@iInitParent2RecordID	int,
				@iElementType		integer, 
				@iTempElementID		integer,
				@sValue				varchar(8000),
				@fValidRecordID		bit,
				@fDeletedValue		bit,
				@iPersonnelTableID	integer,
				@iCount				integer,
				@sSQL				nvarchar(4000),
				@sSQLParam			nvarchar(4000)
		
			SELECT @iWorkflowID = isnull(WE.workflowID, 0),
				@iBaseTableID = isnull(WF.baseTable, 0),
				@piItemType = isnull(WEI.itemType, 0),
				@sElementIdentifier = upper(rtrim(ltrim(isnull(WEI.WFFormIdentifier, '''')))),
				@sItemIdentifier = upper(rtrim(ltrim(isnull(WEI.WFValueIdentifier, '''')))),
				@iDBColumnID = isnull(WEI.DBColumnID, 0),
				@iDBRecord = isnull(WEI.DBRecord, 0)
			FROM ASRSysWorkflowElementItems WEI
			INNER JOIN ASRSysWorkflowElements WE ON WEI.elementID = WE.ID
			INNER JOIN ASRSysWorkflows WF ON WE.workflowID = WF.ID
			WHERE WEI.ID = @piElementItemID
		
			IF @piItemType = 19 -- DB File
			BEGIN
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
					@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
					@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
					@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
		
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
		
				SET @fDeletedValue = 0
		
				SELECT @sTableName = ASRSysTables.tableName, 
					@iRequiredTableID = ASRSysTables.tableID, 
					@sColumnName = ASRSysColumns.columnName,
					@piDBColumnType = ASRSysColumns.dataType,
					@piOLEType = ASRSysColumns.OLEType
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
				END
		
				IF @iDBRecord = 1
				BEGIN
					-- Id'


	SET @sSPCode_1 = 'entified record.
					SELECT @iElementType = ASRSysWorkflowElements.type, 
						@iTempElementID = ASRSysWorkflowElements.ID
					FROM ASRSysWorkflowElements
					WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
						AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sElementIdentifier)))
						
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
							AND Es.identifier = @sElementIdentifier
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
		
					IF @fV'


	SET @sSPCode_2 = 'alidRecordID = 0
					BEGIN
						SET @psErrorMessage = ''Record has been deleted or not selected.''
						RETURN
					END
				END
					
				IF @fDeletedValue = 0
				BEGIN
					IF (@piOLEType = 0) OR (@piOLEType = 1)
					BEGIN
						SET @sSQL = ''SELECT @psFileName = '' + @sTableName + ''.'' + @sColumnName +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
						SET @sSQLParam = N''@psFileName varchar(8000) OUTPUT''
						EXEC sp_executesql @sSQL, @sSQLParam, @psFileName OUTPUT
					END
					ELSE
					BEGIN
						SET @sSQL = ''SELECT '' + @sTableName + ''.'' + @sColumnName + '' AS [file]'' +
							'' FROM '' + @sTableName +
							'' WHERE '' + @sTableName + ''.ID = '' + convert(nvarchar(4000), @iRecordID)
						EXEC sp_executesql @sSQL
					END
				END
			END
			
			IF @piItemType = 20 -- WF File
			BEGIN
				SELECT @iElementID = isnull(ID, 0)
				FROM ASRSysWorkflowElements
				WHERE workflowID = @iWorkflowID
					AND upper(ltrim(rtrim(isnull(identifier, '''')))) = @sElementIdentifier
		
				SELECT @psContentType = fileUpload_contentType,
					@psFileName = fileUpload_fileName
				FROM ASRSysWorkflowInstanceValues
				WHERE instanceID = @piInstanceID
					AND elementID = @iElementID
					AND upper(ltrim(rtrim(isnull(identifier, '''')))) = @sItemIdentifier
		
				SELECT fileUpload_file AS [file]
				FROM ASRSysWorkflowInstanceValues
				WHERE instanceID = @piInstanceID
					AND elementID = @iElementID
					AND upper(ltrim(rtrim(isnull(identifier, '''')))) = @sItemIdentifier
			END
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

/* ------------------------------------------------------------- */
PRINT 'Step 3 of X - Modifying Overnight Job stored procedures'

/* ------------------------------------------------------------- */
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
				SET @sVarCommand = ''UPDATE STATISTICS ['' + @sTableName + ''] WITH FULLSCAN''
				EXEC (@sVarCommand)
				FETCH NEXT FROM tables INTO @sTableName
			END
		
			-- Close and deallocate the cursor
			CLOSE tables
			DEALLOCATE tables
		
		END'

	EXECUTE (@sSPCode_0)



/* ------------------------------------------------------------- */


/* ------------------------------------------------------------- */
PRINT 'Step 3 of X - Modifying Self-service Intranet tables'

/* ------------------------------------------------------------- */


	/* ASRSysSSIIntranetLinks - Add Application File Path column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysSSIntranetLinks', 'U')
	AND name = 'AppFilePath'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
							AppFilePath [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysSSIntranetLinks
							SET ASRSysSSIntranetLinks.AppFilePath = ''''
							WHERE ASRSysSSIntranetLinks.AppFilePath IS NULL'
		EXEC sp_executesql @NVarCommand

	END

	/* ASRSysSSIIntranetLinks - Add Application Parameters column */
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysSSIntranetLinks', 'U')
	AND name = 'AppParameters'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
							AppParameters [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysSSIntranetLinks
							SET ASRSysSSIntranetLinks.AppParameters = ''''
							WHERE ASRSysSSIntranetLinks.AppParameters IS NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* ASRSysSSIIntranetLinks - Add On-screen Document File Path column*/
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysSSIntranetLinks', 'U')
	AND name = 'DocumentFilePath'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
							DocumentFilePath [varchar] (255) NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysSSIntranetLinks
							SET ASRSysSSIntranetLinks.DocumentFilePath = ''''
							WHERE ASRSysSSIntranetLinks.DocumentFilePath IS NULL'
		EXEC sp_executesql @NVarCommand
	END
	
	/* ASRSysSSIIntranetLinks - Add Display Document Hyperlink column*/
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysSSIntranetLinks', 'U')
	AND name = 'DisplayDocumentHyperlink'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
							DisplayDocumentHyperlink [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysSSIntranetLinks
							SET ASRSysSSIntranetLinks.DisplayDocumentHyperlink = 0
							WHERE ASRSysSSIntranetLinks.DisplayDocumentHyperlink IS NULL'
		EXEC sp_executesql @NVarCommand
	END
	
	/* ASRSysSSIIntranetLinks - Add IsSeparator column*/
	SELECT @iRecCount = COUNT(id) FROM syscolumns
	WHERE id = OBJECT_ID('ASRSysSSIntranetLinks', 'U')
	AND name = 'IsSeparator'

	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysSSIntranetLinks ADD 
							IsSeparator [bit] NULL'
		EXEC sp_executesql @NVarCommand

		SET @NVarCommand = 'UPDATE ASRSysSSIntranetLinks
							SET ASRSysSSIntranetLinks.IsSeparator = 0
							WHERE ASRSysSSIntranetLinks.IsSeparator IS NULL'
		EXEC sp_executesql @NVarCommand
	END
		

/* ------------------------------------------------------------- */
PRINT 'Step 4 of X - Rebranding Payroll Interface'
/* ------------------------------------------------------------- */

	DELETE FROM [ASRSysPermissionCategories] 
	WHERE [ASRSysPermissionCategories].[categoryID] = 41
	
	INSERT INTO [ASRSysPermissionCategories] ([ASRSysPermissionCategories].[categoryID], 
											  [ASRSysPermissionCategories].[description], 
											  [ASRSysPermissionCategories].[picture],
											  [ASRSysPermissionCategories].[listOrder], 
											  [ASRSysPermissionCategories].[categoryKey])
	VALUES(41,'Payroll Transfer','',10,'ACCORD')
	
	SELECT @ptrval = TEXTPTR([ASRSysPermissionCategories].[picture]) 
	FROM [ASRSysPermissionCategories] 
	WHERE [ASRSysPermissionCategories].[categoryID] = 41
	
	WRITETEXT [ASRSysPermissionCategories].[picture] @ptrval 0x0000010001001010000000000000680300001600000028000000100000002000000001001800000000004003000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000A3D5A3A3D5A3000000000000000000000000000000000000000000000000000000000000000000000000A3D5A3A3D5A3CCEECCD0EDD0A3D5A3F1F0E6000000000000000000000000000000000000000000000000000000A3D5A3BCE7BCE6FFE5BDE4BDE6FFE5D4F0CFA3D5A3818281000000000000000000000000000000000000000000A3D5A3BDECBDE6FFE5B7E2B795CA94B8DFB8E6FFE5D4F0CF81828116B1EC5FE7FC000000000000000000000000A3D5A3BAF0BAE6FFE5B2E0B295CA94C2E5C295CA94B8E1B2E6FFE58182811D99DA23A5E435D8FD000000000000000000A3D5A3E6FFE5ABDEAB95CA94C6EFC6ECFFECE8F9E895CA94B8E1B281828156BEED3E98DD6FC7D6A3D5A3000000000000A3D5A3AEE2AE95CA94D1F9D1F1FFF1F5FFF5EFFFEFF4FFE795CA94818281C8FCFF56BEED79DFE7D4F0D29FCF9D000000A3D5A395CA94D1FFD1EDFFEE88B48796C695FFFFF9FFFFF9818281B2F0FCFFFDFFC8FCFF79DFE7CCEACBD4F0D29FD09F95CA94D2FFD1D4FFD3E6FFE5CEF1CAE6FFE5FFFFF98182817AEDFFB6FCFFE0FFFFFFFFFF79DFE7E6FFE5CCEACBA2D1A200000095CA9495CA94DDFFCDE6FFE5FFFFF981828138E7FF62E4FF8CEAFFB4F2FFC8FCFF79DFE7D3EDD2E6FFE5A2D1A200000000000000000095CA9495CA9481828101CDFF12D3FF4ADDFF5DDFFF67E3FF96EBFFF6FDF695CA94D3EDD29CCF9C0000000000000000000000009AE5FB35DCFF26D2FF01CBFF06D3FF51E0FFDBFAFDF6FDF6DAE9CBF7FEF495CA949DD19D0000000000000000000000000000000000007BE1FC2CD7FF61DCE9DBFAFDFFFFFFFFFFFFFFFFFF95CA9495CA9400000000000000000000000000000000000000000000000000000000000095CA94E4F3DBF6FDF698CB9800000000000000000000000000000000000000000000000000000000000000000000000000000095CA9495CA94000000000000000000000000FFFF0000F9FF0000E07F0000C03F0000800F0000000700000003000000010000000000000000000080000000E0000000F0000000FC010000FF870000FFCF0000

	-- BatchJob permission category picture incorrectly modified in first draft of 3.7 update script.
	-- Restore original image here.
	UPDATE [ASRSysPermissionCategories] 
	SET [picture] = ''
	WHERE [ASRSysPermissionCategories].[categoryID] = 2

	SELECT @ptrval = TEXTPTR([ASRSysPermissionCategories].[picture]) 
	FROM [ASRSysPermissionCategories] 
	WHERE [ASRSysPermissionCategories].[categoryID] = 2
	
	WRITETEXT [ASRSysPermissionCategories].[picture] @ptrval 0x00000100010010100000010008006805000016000000280000001000000020000000010008000000000000010000000000000000000000010000000100000000000018181800212121006B2929006B313100733131007B4A4200844A4A00845252008C5252008C5A5A00945A5A009C636300946B63009C6B63009C6B6B009C736B009C737300A56B6B00A5737300A57B7300A57B7B00AD7B7B00AD7B840094949400A5848400AD848400A58C8400AD8C8400B5848400B58C8C00AD949400B5949400BD949400B59C9400BD9C9C00BDA59C00A5A5A500BDADA500C6949C00C69C9C00C6A5A500CEA5A500C6ADA500C6ADAD00CEADAD00D6ADAD00CEB5B500D6B5B500DEB5B500D6BDB500D6BDBD00C6C6C600CECECE00DEC6C600DECECE00D6D6D600DED6D600DEDEDE00E7CEC600E7CED600F7D6D600F7DEDE00F7DEE700EFE7E700EFEFEF00F7E7E700FFE7E700FFEFEF00FFF7F700FFF7FF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFFFF000000000000200F070D1C260000000000000000230C13131D16130A0B000000000000210F1E333F3C3D3127120F000000002D0F2139FF45003E431F210F140000001315420025FF4545440029270B0000320E2CFFFFFFFFFF4545434331131B002D094041FFFFFFFFFF46453E3D1A0D002806FF0038FFFF1834FF4500371D0800300740FFFFFFFF350134FF463F161100360C2BFF38FFFFFF350134FF33132200001A0F400041FF3AFF3502401E0F0000003B0F1940FFFF0038FF40210F240000000F040F0F2C40FF402F150F05040000000F2A04170C0707090F13042E0400000000100F003630042D33001A1A00000000000000000007040300000000000000F81F0000E00F0000C00700008003000080030000000100000001000000010000000100000001000080030000800300008003000080030000C8270000FC7F0000


/* ------------------------------------------------------------- */
PRINT 'Step 5 of X - Updating sp_ASRGetOrderDefinition'

/* ------------------------------------------------------------- */

	----------------------------------------------------------------------
	-- sp_ASRGetOrderDefinition
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRGetOrderDefinition]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRGetOrderDefinition]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[sp_ASRGetOrderDefinition]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE sp_ASRGetOrderDefinition (
			@piOrderID int) 
		AS
		BEGIN
			/* Return the recordset of order items for the given order. */
			SELECT ASRSysOrderItems.*,
				ASRSysColumns.columnName,
				ASRSysColumns.tableID,
				ASRSysColumns.dataType,
			    	ASRSysTables.tableName,
					ASRSysColumns.Size,
					ASRSysColumns.Decimals,
					ASRSysColumns.Use1000Separator, 
					ASRSysColumns.blankIfZero
			FROM ASRSysOrderItems
			INNER JOIN ASRSysColumns 
				ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
			INNER JOIN ASRSysTables 
				ON ASRSysTables.tableID = ASRSysColumns.tableID
			WHERE ASRSysOrderItems.orderID = @piOrderID
			ORDER BY ASRSysOrderItems.type, 
				ASRSysOrderItems.sequence
		END'

	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 6 of X - Updating spASRAccordPopulateTransaction'
/* ------------------------------------------------------------- */

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRAccordPopulateTransaction]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRAccordPopulateTransaction]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRAccordPopulateTransaction] (
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
			DECLARE @bForceAsUpdate bit
		
			SET @piTransactionID = null
			SET @bCreate = 1
			SET @bForceAsUpdate = 0
		
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
						AND TransferType = @piTransferType
					ORDER BY CreatedDateTime DESC
				
					IF @iStatus IS NULL OR @iStatus IN (20, 23)
					BEGIN
						SET @piTransactionType = 0
						SET @pbSendAllFields = 1
					END
				END
		
				SELECT @bForceAsUpdate = ForceAsUpdate FROM ASRSysAccordTransferTypes
				WHERE TransferTypeID = @piTransferType
		
				IF @bForceAsUpdate = 1 AND @piTransactionType = 0 SET @piTransactionType = 1
		
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

	EXECUTE (@sSPCode_0)

/* --------------------------------------------------------------------- */
PRINT 'Step 7 of X - Updating ASRSysAccordTransferFieldDefinitions table'
/* --------------------------------------------------------------------- */

  IF (SELECT COUNT(*) FROM sysobjects WHERE xtype='u' AND name = 'ASRSysAccordTransferFieldDefinitions' ) <> 0
  BEGIN
	  UPDATE ASRSysAccordTransferFieldDefinitions
	  SET    AlwaysTransfer = 1
	  WHERE  IsKeyField = 1
	     AND AlwaysTransfer = 0
	     AND TransferTypeID IN
		     (SELECT TransferTypeID 
				  FROM ASRSysAccordTransferTypes
				  WHERE  TransferType LIKE 'Extra Deduction - User Defined %')
  END	


/* ------------------------------------------------------------- */
PRINT 'Step 8 of X - Updating Support Email Address'
/* ------------------------------------------------------------- */

  SET @sSPCode_0 = 'DELETE FROM [ASRSysSystemSettings] 
              WHERE [ASRSysSystemSettings].[Section] = ''support'' 
                AND [ASRSysSystemSettings].[SettingKey] = ''email'''
	EXECUTE (@sSPCode_0)
	
  SET @sSPCode_0 = 'INSERT INTO [ASRSysSystemSettings] ([ASRSysSystemSettings].[Section], 
                                                  [ASRSysSystemSettings].[SettingKey], 
                                                  [ASRSysSystemSettings].[SettingValue])
                    VALUES (''support'', ''email'', ''HCMsupport@coasolutions.com'')'
              
	EXECUTE (@sSPCode_0)



/* ---------------------------------------------------------------------- */
PRINT 'Step 9 of X - Updating spASREmailImmediate'
/* ---------------------------------------------------------------------- */

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASREmailImmediate]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASREmailImmediate]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASREmailImmediate]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].[spASREmailImmediate](@Username varchar(255)) AS
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
								@TableID int,
								@tmpUser varchar(100)
							DECLARE @RecipTo varchar(4000),
								@TempText nvarchar(4000),
								@CC varchar(4000),
								@BCC varchar(4000),
								@Subject varchar(4000),
								@MsgText varchar(8000),
								@Attachment varchar(4000)
				
							DECLARE emailqueue_cursor
							CURSOR LOCAL FAST_FORWARD FOR 
							  SELECT QueueID
							       , ASRSysEmailQueue.LinkID
							       , RecordID
							       , ASRSysEmailQueue.ColumnID
							       , ColumnValue
							       , RecordDesc
							       , RecalculateRecordDesc
							       , ASRSysEmailQueue.TableID
							       , DateDue
							       , UserName
							  FROM   ASRSysEmailQueue
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
							FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc, @RecalculateRecordDesc, @TableID, @DateDue, @tmpUser
				
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
												
												if ltrim(rtrim(@Username)) <> '''' and @Username is not null EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, @Username, @RecipTo OUTPUT, @CC OUTPUT, @BCC OUTPUT, @Subject OUTPUT, @MsgText OUTPUT, @Attachment OUTPUT
												Else EXEC @hResult = @sSQL @recordid, @recDesc, @columnvalue, @emailDate, @tmpUser, @RecipTo OUTPUT, @CC OUTPUT, @BCC OUTPUT, @Subject OUTPUT, @MsgText OUTPUT, @Attachment OUTPUT
												
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
												EXEC spASRSendMail @hResult'


	SET @sSPCode_1 = ' OUTPUT, @RecipTo, '''', '''', @Subject,  @MsgText, ''''
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
								FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @ColumnID, @ColumnValue, @RecDesc, @RecalculateRecordDesc, @TableID, @DateDue, @tmpUser
							END
							CLOSE emailqueue_cursor
							DEALLOCATE emailqueue_cursor
						END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

/* ---------------------------------------------------------------------- */
PRINT 'Step 10 of X - Updating spASRGetActualUserDetails'
/* ---------------------------------------------------------------------- */

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
				@piUserGroupID integer OUTPUT,
				@piModuleKey varchar(20)
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
					AND CASE 
						WHEN (usg.uid IS null) THEN null
						ELSE usg.name
						END IN (
									SELECT [groupName]
									FROM dbo.[ASRSysGroupPermissions]
									WHERE itemID IN (
																		SELECT [itemID]
																		FROM dbo.[ASRSysPermissionItems]
																		WHERE categoryID = 1
																		AND itemKey LIKE @piModuleKey + ''%''
																	)  
									AND [permitted] = 1
			)
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
					AND CASE 
						WHEN (usg.uid IS null) THEN null
						ELSE usg.name
						END IN (
									SELECT [groupName]
									FROM dbo.[ASRSysGroupPermissions]
									WHERE itemID IN (
																		SELECT [itemID]
																		FROM dbo.[ASRSysPermissionItems]
																		WHERE categoryID = 1
																		AND itemKey LIKE @piModuleKey + ''%''
																	)  
									AND [permitted] = 1
			)
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

/* ------------------------------------------------------------- */
PRINT 'Step 11 of X - Updating Log Purging'

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[INS_AsrSysPurgeEventLog]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[INS_ASRSysPurgeEventLog]

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[INS_AsrSysPurgeAuditTrail]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[INS_AsrSysPurgeAuditTrail]

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[INS_AsrSysPurgeAuditGroup]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[INS_AsrSysPurgeAuditGroup]

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[INS_AsrSysPurgeAuditPermissions]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[INS_AsrSysPurgeAuditPermissions]

		if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[INS_AsrSysPurgeAuditAccess]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
		drop trigger [dbo].[INS_AsrSysPurgeAuditAccess]


/* ------------------------------------------------------------- */
PRINT 'Step 12 of X - Updating Support Web Site Address'
/* ------------------------------------------------------------- */

  SET @sSPCode_0 = 'DELETE FROM [ASRSysSystemSettings] 
              WHERE [ASRSysSystemSettings].[Section] = ''support'' 
                AND [ASRSysSystemSettings].[SettingKey] = ''webpage'''
	EXECUTE (@sSPCode_0)
	
  SET @sSPCode_0 = 'INSERT INTO [ASRSysSystemSettings] ([ASRSysSystemSettings].[Section], 
                                                  [ASRSysSystemSettings].[SettingKey], 
                                                  [ASRSysSystemSettings].[SettingValue])
                    VALUES (''support'', ''webpage'', ''www.mycoasolutions.com'')'
              
	EXECUTE (@sSPCode_0)




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
values('database', 'version', '3.7')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '3.7.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.4.0')

delete from asrsyssystemsettings
where [Section] = '.NET Assembly' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('.NET Assembly', 'minimum version', '3.7.0')

delete from asrsyssystemsettings
where [Section] = 'outlook service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('outlook service', 'minimum version', '3.6.0')

delete from asrsyssystemsettings
where [Section] = 'workflow service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('workflow service', 'minimum version', '3.7.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v3.7')


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
	GRANT EXECUTE ON xp_EnumGroups TO public';
EXEC sp_executesql @NVarCommand;

-- Version specific functions
IF (@iSQLVersion < 11)
BEGIN
	SELECT @NVarCommand = 'USE master
		GRANT EXECUTE ON xp_StartMail TO public
		GRANT EXECUTE ON xp_SendMail TO public';
	EXEC sp_executesql @NVarCommand;
END



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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v3.7 Of HR Pro'
