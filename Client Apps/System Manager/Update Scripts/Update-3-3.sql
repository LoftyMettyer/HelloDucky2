
/* --------------------------------------------------- */
/* Update the database from version 3.2 to version 3.3 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(8000),
	@iSQLVersion numeric(3,1),
	@NVarCommand nvarchar(4000)

DECLARE @sSQL varchar(8000)
DECLARE @sSPCode_0 nvarchar(4000)
DECLARE @sSPCode_1 nvarchar(4000)
DECLARE @sSPCode_2 nvarchar(4000)
DECLARE @sSPCode_3 nvarchar(4000)
DECLARE @sSPCode_4 nvarchar(4000)
DECLARE @sSPCode_5 nvarchar(4000)
DECLARE @sSPCode_6 nvarchar(4000)
DECLARE @sSPCode_7 nvarchar(4000)

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

/* Exit if the database is not version 3.2 or 3.3. */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '3.2') and (@sDBVersion <> '3.3')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
PRINT 'Step 1 of 16 - Creating/modifying Workflow tables'

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowStepDelegation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowStepDelegation] (
						[ID] [int] IDENTITY (1, 1) NOT NULL ,
						[StepID] [int] NOT NULL ,
						[DelegateEmail] [varchar] (8000) NULL 
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	/* Add new timeout frequency column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'TimeoutFrequency'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						TimeoutFrequency [int] NULL'
		EXEC sp_executesql @NVarCommand

	END

	/* Add new timeout period column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElements')
	and name = 'TimeoutPeriod'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElements ADD 
						TimeoutPeriod [int] NULL'
		EXEC sp_executesql @NVarCommand

	END

	/* Add new Lookup Table ID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'LookupTableID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						LookupTableID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	/* Add new Lookup Column ID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowElementItems')
	and name = 'LookupColumnID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowElementItems ADD 
						LookupColumnID [int] NULL'
		EXEC sp_executesql @NVarCommand
	END
	
	/* Create Workflow Triggered Link Columns Table */
	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowTriggeredLinkColumns]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowTriggeredLinkColumns](
						[linkID] [int] NULL,
						[columnID] [int] NULL
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	/* Create Workflow Triggered Links Table */
	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowTriggeredLinks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowTriggeredLinks](
						[linkID] [int] NULL,
						[workflowID] [int] NULL,
						[tableID] [int] NULL,
						[filterID] [int] NULL,
						[effectiveDate] [datetime] NULL,
						[type] [smallint] NULL,
						[recordInsert] [bit] NULL,
						[recordUpdate] [bit] NULL,
						[recordDelete] [bit] NULL,
						[dateColumn] [int] NULL,
						[dateOffset] [int] NULL,
						[dateOffsetPeriod] [smallint] NULL
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowElementItemValues]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowElementItemValues](
									[itemID] [int] NOT NULL,
									[value] [varchar](255) NOT NULL,
									[sequence] [int] NOT NULL
								) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand

		SELECT @NVarCommand = '	CREATE TRIGGER [dbo].[DEL_ASRSysWorkflowElementItemValues] 
					   ON  [dbo].[ASRSysWorkflowElementItems] 
					   FOR DELETE
					AS 
					BEGIN
						DELETE FROM ASRSysWorkflowElementItemValues WHERE itemID NOT IN (SELECT ID FROM ASRSysWorkflowElementItems)
					END'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflows')
	and name = 'initiationType'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflows ADD 
						initiationType [smallint] NULL'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflows')
	and name = 'baseTable'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflows ADD 
						baseTable [int] NULL'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowQueue]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowQueue](
						[QueueID] [int] IDENTITY(1,1) NOT NULL,
						[LinkID] [int] NULL,
						[Immediate] [bit] NULL,
						[RecordID] [int] NULL,
						[DateDue] [datetime] NULL,
						[DateInitiated] [datetime] NULL,
						[RecordDesc] [varchar](255),
						[UserName] [varchar](50),
						[RecalculateRecordDesc] [bit] NULL
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	if not exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ASRSysWorkflowQueueColumns]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	BEGIN
		SELECT @NVarCommand = 'CREATE TABLE [dbo].[ASRSysWorkflowQueueColumns](
						[QueueID] [int] NULL,
						[ColumnID] [int] NULL,
						[ColumnValue] [varchar](8000) NULL
					) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

	/* Add new InstanceValues ColumnID column */
	SELECT @iRecCount = count(id) FROM syscolumns
	where id = (select id from sysobjects where name = 'ASRSysWorkflowInstanceValues')
	and name = 'ColumnID'

	if @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'ALTER TABLE ASRSysWorkflowInstanceValues ADD 
						ColumnID [int] NULL'
		EXEC sp_executesql @NVarCommand

	END

/* ------------------------------------------------------------- */
PRINT 'Step 2 of 16 - Modifying Workflow triggers'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DEL_ASRSysWorkflowInstanceSteps]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
	drop trigger [dbo].[DEL_ASRSysWorkflowInstanceSteps]

	SELECT @NVarCommand = 'CREATE TRIGGER DEL_ASRSysWorkflowInstanceSteps ON [dbo].[ASRSysWorkflowInstanceSteps] 
		FOR DELETE
		AS
		BEGIN
			/* Delete related records. */
			DELETE FROM ASRSysWorkflowStepDelegation
			WHERE ASRSysWorkflowStepDelegation.stepID IN (SELECT id FROM deleted)
		END'

	EXEC sp_executesql @NVarCommand

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DEL_ASRSysWorkflowTriggeredLinks]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
	drop trigger [dbo].[DEL_ASRSysWorkflowTriggeredLinks]

	SELECT @NVarCommand = '	CREATE TRIGGER [dbo].[DEL_ASRSysWorkflowTriggeredLinks] 
				   ON  [dbo].[ASRSysWorkflowTriggeredLinks] 
				   FOR DELETE
				AS 
				BEGIN
					DELETE FROM ASRSysWorkflowTriggeredLinkColumns WHERE LinkID NOT IN (SELECT LinkID FROM ASRSysWorkflowTriggeredLinks)
				END'
	EXEC sp_executesql @NVarCommand

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DEL_ASRSysWorkflowQueue]') and OBJECTPROPERTY(id, N'IsTrigger') = 1)
	drop trigger [dbo].[DEL_ASRSysWorkflowQueue]

	SELECT @NVarCommand = 'CREATE TRIGGER [dbo].[DEL_ASRSysWorkflowQueue] ON [dbo].[ASRSysWorkflowQueue] 
		FOR DELETE
		AS
		BEGIN
			/* Delete related records. */
			DELETE FROM ASRSysWorkflowQueueColumns
			WHERE ASRSysWorkflowQueueColumns.queueID IN (SELECT queueID FROM deleted)
		END'

	EXEC sp_executesql @NVarCommand

/* ------------------------------------------------------------- */
PRINT 'Step 3 of 16 - Modifying Workflow stored procedures'

	IF NOT EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetWorkflowDelegatedRecords]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
	BEGIN
		SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetWorkflowDelegatedRecords]
			(
			    @psOriginalRecipient varchar(8000),
			    @pcurDelegatedRecords cursor varying output
			)
			AS
			BEGIN
			    SET @pcurDelegatedRecords = CURSOR FORWARD_ONLY STATIC FOR
			        SELECT 0 AS [ID]
			    OPEN @pcurDelegatedRecords
			END'
		EXECUTE (@sSPCode_0)
	END

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

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRWorkflowValidRecord
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
				SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_TablePersonnel''
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
		
			SELECT @sTableName = isnull(ASRSysTables.tableName, '''')
			FROM ASRSysTables
			WHERE ASRSysTables.tableID = @iTableID
		
			IF len(@sTableName) > 0 
			BEGIN
				SET @sSQL = ''SELECT @iRecCount = COUNT(*) FROM '' + @sTableName + '' WHERE ID = '' + convert(nvarchar(4000), @piRecordID)
				SET @sParam = N''@iRecCount integer OUTPUT''
				EXEC sp_executesql @sSQL, @sParam, @iRecCount OUTPUT
		
				IF @iRecCount > 0 SET @pfValid = 1
			END	
		END'

	EXECUTE (@sSPCode_0)

	----------------------------------------------------------------------
	-- spASRDelegateWorkflowEmail
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRDelegateWorkflowEmail]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRDelegateWorkflowEmail]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRDelegateWorkflowEmail]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE spASRDelegateWorkflowEmail 
		(
			@psTo			varchar(8000),
			@psMessage		varchar(8000),
			@piStepID		integer,
			@psEmailSubject	varchar(8000),
			@piLevel		integer
		)
		AS
		BEGIN
			DECLARE
				@iDelegateEmailID	integer,
				@iCount		integer,
				@iTemp		integer,
				@sTemp		varchar(8000),
				@fDelegate		bit,
				@curDelegatedRecords	cursor,
				@fDelegationValid	bit,
				@sAllDelegateTo	varchar(8000),
				@iDelegateRecordID	integer,
				@sDelegateTo		varchar(8000),
				@sDelegatedMessage	varchar(8000),
				@iCurrentStepID 	integer,
				@fCopyDelegateEmail	bit,
				@sSQL			nvarchar(4000),
				@iInstanceID		integer
		
			-- Get the instanceID of the given step
			SELECT @iInstanceID = instanceID
			FROM ASRSysWorkflowInstanceSteps
			WHERE ID = @piStepID
		
			-- Get the delegate email address definition. 
			SET @iDelegateEmailID = 0
			SELECT @sTemp = ISNULL(parameterValue, '''')
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_DelegateEmail''
			SET @iDelegateEmailID = convert(integer, @sTemp)
		
			SET @fCopyDelegateEmail = 1
			SELECT @sTemp = LTRIM(RTRIM(UPPER(ISNULL(parameterValue, ''''))))
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_CopyDelegateEmail''
			IF @sTemp = ''FALSE''
			BEGIN
				SET @fCopyDelegateEmail = 0
			END
		
			CREATE TABLE #delegatedRecords (ID integer,
				email	varchar(8000))
		
			SET @fDelegate = 0 -- Flag whether or not any delegation is required.
		
			-- Clear out the delegation record for the current step (only do this for this stored procedure''s first call0
			IF @piLevel = 0 
			BEGIN
				DELETE FROM ASRSysWorkflowStepDelegation
				WHERE stepID = @piStepID
			END
			SET @piLevel = @piLevel + 1
		
			IF @iDelegateEmailID > 0 
			BEGIN
				-- Get the ID of the personnel records to which the original recipient is delegating to.
				EXEC spASRGetWorkflowDelegatedRecords @psTo, @curDelegatedRecords OUTPUT
		
				FETCH NEXT FROM @curDelegatedRecords INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #delegatedRecords (ID, email) 
					VALUES (@iTemp, '''')
		
					SET @fDelegate = 1 -- Flag to indicate that delegation is required.
					
					FETCH NEXT FROM @curDelegatedRecords INTO @iTemp 
				END
				CLOSE @curDelegatedRecords
				DEALLOCATE @curDelegatedRecords
			END
		
			-- Delegation IS required, so find out the delegation email addresses
			IF @fDelegate = 1
			BEGIN
				SET @fDelegationValid = 0 -- Flag whether or not delegation is valid. ie. if we could determine any email addresses for the delegates
				SET @sAllDelegateTo = ''''
		
				-- Calculate the delegated email addresses
				DECLARE delegatesCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ID
				FROM #delegatedRecords
		
				OPEN delegatesCursor
				FETCH NEXT FROM delegatesCursor INTO @iDelegateRecordID
				WHILE (@@fetch_status = 0)
				BEGIN
					SET @sDelegateTo = ''''
					SET @sSQL = ''spASRSysEmailAddr''
			
					IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
					BEGIN
						-- Get the delegate''s email address
						EXEC @sSQL @sDelegateTo OUTPUT, @iDelegateEmailID, @iDelegateRecordID
						IF @sDelegateTo IS null SET @sDelegateTo = ''''
					END
		
					IF len(@sDelegateTo) > 0 
					BEGIN
						-- Check if this step has already been sent to the delegate.
						-- If so, we do NOT want to do it again.
						SELECT @iCount = COUNT(*)
						FROM ASRSysWorkflowInstanceSteps
						WHERE userEmail = @sDelegateTo
							AND ID = @piStepID
		
						IF @iCount > 0 
						BEGIN
							SET @sDelegateTo = ''''
						END
						ELSE
						BEGIN
							SELECT @iCount = COUNT(*)
							FROM ASRSysWorkflowStepDelegation
							WHERE delegateEmail = @sDelegateTo
								AND stepID = @piStepID
			
							IF @iCount > 0 SET @sDelegateTo = '''''


	SET @sSPCode_1 = '
						END
					END
						
					IF len(@sDelegateTo) > 0 
					BEGIN
						SET @fDelegationValid = 1 -- Email address has been determined so delegation IS valid
						SET @sAllDelegateTo = @sAllDelegateTo + 
							CASE 
								WHEN len(@sAllDelegateTo) > 0 THEN ''; ''
								ELSE ''''
							END + @sDelegateTo
		
						UPDATE #delegatedRecords 
						SET email = @sDelegateTo
						WHERE ID = @iDelegateRecordID
			
						INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
						VALUES (@sDelegateTo, @piStepID)
		
						exec spASRDelegateWorkflowEmail 
							@sDelegateTo,
							@psMessage,
							@piStepID,
							@psEmailSubject,
							@piLevel
					END
					
					FETCH NEXT FROM delegatesCursor INTO @iDelegateRecordID
				END
				CLOSE delegatesCursor
				DEALLOCATE delegatesCursor
		
				/* No delegate emails were determined, so do not delegate. */
				IF @fDelegationValid = 0 SET @fDelegate = 0
			END 
		
			IF @fDelegate = 1
			BEGIN
				-- Delegation IS required, AND IS available, so copy the email to the original recipient (if required)
				IF @fCopyDelegateEmail = 1
				BEGIN
					SET @sDelegatedMessage = ''The following email has been delegated to '' + @sAllDelegateTo + char(13) + 
						''--------------------------------------------------'' + char(13) +
						@psMessage 
			
					-- Send the email. 
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
						Subject)
					VALUES ('''',
						'''',
						getdate(),
						''HR Pro Workflow'',
						1,
						0, 
						@psTo,
						@sDelegatedMessage,
						@iInstanceID,
						@psEmailSubject)
				END
			END
			ELSE
			BEGIN
				-- No delegation required, or available, so send the email to the original recipient
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
					Subject)
				VALUES ('''',
					'''',
					getdate(),
					''HR Pro Workflow'',
					1,
					0, 
					@psTo,
					@psMessage,
					@iInstanceID,
					@psEmailSubject)
			END
		
			DROP TABLE #delegatedRecords
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

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
				@curDelegatedRecords	cursor,
				@fDelegate		bit,
				@fDelegationValid	bit,
				@fCopyDelegateEmail	bit,
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
				@sEmailSubject	varchar(200)

			SELECT @iCurrentStepID = ID
			FROM ASRSysWorkflowInstanceSteps
			WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
				AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

			SET @fCopyDelegateEmail = 1
			SELECT @sTemp = LTRIM(RTRIM(UPPER(ISNULL(parameterValue, ''''))))
			FROM ASRSysModuleSetup
			WHERE moduleKey = ''MODULE_WORKFLOW''
				AND parameterKey = ''Param_CopyDelegateEmail''
			IF @sTemp = ''FALSE''
			BEGIN
				SET @fCopyDelegateEmail = 0
			END

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
					ASRSysWorkflowInstanceValues.valueDescription = @sValueDesc'


	SET @sSPCode_1 = 'ription
				WHERE ASRSysWorkflowInstanceValues.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceValues.elementID = @piElementID
					AND isnull(ASRSysWorkflowInstanceValues.columnID, 0) = 0
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
					SET @sValue = left(@sValue, 1000)

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
						SELECT @iRecDescID =  ASRSysTables.RecordDescExprID
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
				END
			END
					
			SET @hResult = 0
			SET @sTo = ''''
		
			IF @iElementType = 3 -- Email element
			BEGIN
				-- Get the email recipient. 
				SET @sTo = ''''
				SET @iEmailRecordID = 0
				SET @sSQL = ''spASRSysEmailAddr''

				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					-- Get the record ID required. 
					IF (@iEmailRecord = 0) OR (@iEmailRecord = 4)
					BEGIN
						/* Initiator record. */
						SELECT @iEmailRecordID = ASRSysWorkflowInstances.initiatorID
						FROM ASRSysWorkflowInstances
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID
					END
		
		'


	SET @sSPCode_2 = '			IF @iEmailRecord = 1
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
								AND IV.identifier = Es.identifier
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sRecSelWebFormIdentifier
							WHERE IV.instanceID = @piInstanceID
						END
					END

					SET @fValidRecordID = 1
					IF (@iEmailRecord = 0) OR (@iEmailRecord = 1) OR (@iEmailRecord = 4)
					BEGIN
						EXEC spASRWorkflowValidRecord
							@piInstanceID,
							@iEmailRecord,
							@iEmailRecordID,
							@sRecSelWebFormIdentifier,
							@sRecSelIdentifier,
							@fValidRecordID	OUTPUT

						IF @fValidRecordID = 0
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Email record has been deleted.''
										
							SET @hResult = -1
						END
					END

					IF @fValidRecordID = 1
					BEGIN
						/* Get the recipient address. */
						EXEC @hResult = @sSQL @sTo OUTPUT, @iEmailID, @iEmailRecordID
						IF @sTo IS null SET @sTo = ''''

						IF LEN(rtrim(ltrim(@sTo))) = 0
						BEGIN
							-- Email step failure if no known recipient.
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''No email recipient.''
										
							SET @hResult = -1
						END
					END
				END
		
				IF LEN(rtrim(ltrim(@sTo))) > 0
				BEGIN
					IF (rtrim(ltrim(@sTo)) = ''@'')
						OR (charindex('' @ '', @sTo) > 0)
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.userEmail = @sTo
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

						EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Invalid email recipient.''
						
						SET @hResult = -1
					END
					ELSE
					BEGIN
						/* Build the email message. */
						EXEC spASRGetWorkflowEmailMessage @piInstanceID, @piElementID, @sMessage OUTPUT, @fValidRecordID OUTPUT
		
						IF @fValidRecordID = 1
						BEGIN
							exec spASRDelegateWorkflowEmail 
								@sTo,
								@sMessage,
								@iCurrentStepID,
								@sEmailSubject,
								0
						END
						ELSE
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Email item database value record has been deleted.''
										
							SET @hResult = -1
						END
					END
				END
			END
		
			IF @hResult = 0
			BEGIN
				/* Update the ASRSysWorkflowInstanceSteps table to show that this step has completed, and the next step(s) are now activated. */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 3,
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate(),
					ASRSysWorkflow'


	SET @sSPCode_3 = 'InstanceSteps.userEmail = CASE
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
		
					IF @iElementType = 2 -- WebForm
					BEGIN
						EXEC spASRGetDecisionSucceedingWorkflowElements @piElementID, 0, @superCursor OUTPUT
					END
					ELSE
					BEGIN
						EXEC spASRGetSucceedingWorkflowElements @piElementID, @superCursor OUTPUT
					END
		
					FETCH NEXT FROM @superCursor INTO @iTemp
					WHILE (@@fetch_status = 0)
					BEGIN
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
							AND WIS.elementID '


	SET @sSPCode_4 = '= @piElementID

						-- Return a list of the workflow form elements that may need to be displayed to the initiator straight away 
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
						DELETE FROM ASRSysWorkflowStepDelegation
						WHERE stepID IN (SELECT ASRSysWorkflowInstanceSteps.ID 
							FROM ASRSysWorkflowInstanceSteps
							WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
								AND ASRSysWorkflowInstanceSteps.elementID IN 
									(SELECT #succeedingElements.elementID
									FROM #succeedingElements)
								AND ASRSysWorkflowInstanceSteps.status = 0)
						
						INSERT INTO ASRSysWorkflowStepDelegation (delegateEmail, stepID)
						(SELECT WSD.delegateEmail,
							SuccWIS.ID
						FROM ASRSysWorkflowStepDelegation WSD
						INNER JOIN ASRSysWorkflowInstanceSteps CurrWIS ON WSD.stepID = CurrWIS.ID
						INNER JOIN ASRSysWorkflowInstanceSteps SuccWIS ON CurrWIS.instanceID = SuccWIS.instanceID
							AND SuccWIS.elementID IN (SELECT #succeedingElements.elementID
								FROM #succeedingElements)
							AND SuccWIS.status = 0
						INNER JOIN ASRSysWorkflowElements SuccWE ON SuccWIS.elementID = SuccWE.ID
							AND SuccWE.type = 2
						WHERE WSD.stepID = @iCurrentStepID)

						UPDATE ASRSysWorkflowInstanceSteps
						SET ASRSysWorkflowInstanceSteps.status = 1,
							ASRSysWorkflowInstanceSteps.activationDateTime = getdate(),
							ASRSysWorkflowInstanceSteps.userEmail = CASE
								WHEN (SELECT ASRSysWorkflowElements.type 
									FROM ASRSysWorkflowElements 
									WHERE ASRSysWorkflowElements.id = ASRSysWorkflowInstanceSteps.elementID) = 2 THEN @sTo -- 2 = Web Form element
								ELSE ASR'


	SET @sSPCode_5 = 'SysWorkflowInstanceSteps.userEmail
							END
						WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
							AND ASRSysWorkflowInstanceSteps.elementID IN 
								(SELECT #succeedingElements.elementID
								FROM #succeedingElements)
							AND ASRSysWorkflowInstanceSteps.status = 0
					END
					
					DROP TABLE #succeedingElements
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
					ASRSysWorkflowInstanceSteps.completionDateTime = getdate()
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
					
					-- NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					-- is performed by a DELETE trigger on the ASRSysWorkflowInstances table. 
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
		+ @sSPCode_4
		+ @sSPCode_5)

	----------------------------------------------------------------------
	-- spASRWorkflowOutOfOfficeConfigured
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowOutOfOfficeConfigured]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowOutOfOfficeConfigured]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowOutOfOfficeConfigured]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE dbo.spASRWorkflowOutOfOfficeConfigured
		(
		    @pfOutOfOfficeConfigured bit output
		)
		AS
		BEGIN
			DECLARE
				@iCount	integer
		
			-- Check if the SP that checks if the current user is OutOfOffice exists
			SELECT @iCount = COUNT(*)
			FROM sysobjects
			WHERE id = object_id(''spASRWorkflowOutOfOfficeCheck'')
				AND sysstat & 0xf = 4
		
			IF @iCount > 0 
			BEGIN
				-- Check if the SP that sets/resets the current user to be OutOfOffice exists
				SELECT @iCount = COUNT(*)
				FROM sysobjects
				WHERE id = object_id(''spASRWorkflowOutOfOfficeSet'')
					AND sysstat & 0xf = 4
			END
		
			SET @pfOutOfOfficeConfigured = 
			CASE	
				WHEN @iCount > 0 THEN 1
				ELSE 0
			END
		END'

	EXECUTE (@sSPCode_0)

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

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRActionActiveWorkflowSteps]
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
					DEALLOCATE @superCursor'
		
	SET @sSPCode_1 = '
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
				SET ASRSysWorkflowInstanceSteps.status = 6 -- Timeout
				WHERE ASRSysWorkflowInstanceSteps.ID = @iStepID

				-- Activate the succeeding elements on the Timeout flow
				CREATE TABLE #succeedingElements3 (elementID integer)
					
				EXEC spASRGetDecisionSucceedingWorkflowElements @iElementID, 1, @superCursor OUTPUT'
					
	SET @sSPCode_2 = '
				FETCH NEXT FROM @superCursor INTO @iTemp
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #succeedingElements3 (elementID) VALUES (@iTemp)
									
					FETCH NEXT FROM @superCursor INTO @iTemp 
				END
				CLOSE @superCursor
				DEALLOCATE @superCursor
					
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 1,
					ASRSysWorkflowInstanceSteps.activationDateTime = getdate()
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID IN 
						(SELECT #succeedingElements3.elementID 
						FROM #succeedingElements3)
					AND ASRSysWorkflowInstanceSteps.status = 0
					
				DROP TABLE #succeedingElements3

				/* Set activated Web Forms to be ''pending'' (to be done by the user) */
				UPDATE ASRSysWorkflowInstanceSteps
				SET ASRSysWorkflowInstanceSteps.status = 2
				WHERE ASRSysWorkflowInstanceSteps.id IN (
					SELECT ASRSysWorkflowInstanceSteps.ID
					FROM ASRSysWorkflowInstanceSteps
					INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowInstanceSteps.elementID = ASRSysWorkflowElements.ID
					WHERE ASRSysWorkflowInstanceSteps.status = 1
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
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @iInstanceID
					AND ASRSysWorkflowInstanceSteps.status = 3
					AND ASRSysWorkflowElements.type = 1
										
				IF @iCount > 0 
				BEGIN
					UPDATE ASRSysWorkflowInstances
					SET ASRSysWorkflowInstances.completionDateTime = getdate(), 
						ASRSysWorkflowInstances.status = 3
					WHERE ASRSysWorkflowInstances.ID = @iInstanceID
					
					/* NB. Deletion of records in related tables (eg. ASRSysWorkflowInstanceSteps and ASRSysWorkflowInstanceValues)
					is performed by a DELETE trigger on the ASRSysWorkflowInstances table. */
				END

				FETCH NEXT FROM timeoutCursor INTO @iInstanceID, @iElementID, @iStepID
			END
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

	----------------------------------------------------------------------
	-- spASRGetDecisionSucceedingWorkflowElements
	----------------------------------------------------------------------
	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRGetDecisionSucceedingWorkflowElements]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetDecisionSucceedingWorkflowElements]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRGetDecisionSucceedingWorkflowElements]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].[spASRGetDecisionSucceedingWorkflowElements]
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
				AND ((L.startOutboundFlowCode = @piValue) OR 
					(@piValue = 0 and L.startOutboundFlowCode = -1))
				AND E.type <> 8 -- 8 = Connector 1
		
			DECLARE succeedingConnectorsCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT E.connectionPairID
			FROM ASRSysWorkflowLinks L
			INNER JOIN ASRSysWorkflowElements E ON L.endElementID = E.ID
			WHERE L.startElementID = @piElementID
				AND ((L.startOutboundFlowCode = @piValue) OR 
					(@piValue = 0 and L.startOutboundFlowCode = -1))
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
		END'

	EXECUTE (@sSPCode_0)

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

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetWorkflowEmailMessage
					(
						@piInstanceID		integer,
						@piElementID		integer,
						@psMessage		varchar(8000)	OUTPUT, 
						@pfOK	bit	OUTPUT
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
							@iWorkflowID		integer, 
							@fValidRecordID	bit,
							@iColumnID			integer
									
						SET @pfOK = 1
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
							SET @sValue = ''''

							IF @iItemType = 1
							BEGIN
								/* Database value. */
								SELECT @sTableName = ASRSysTables.tableName, 
									@sColumnName = ASRSysColumns.columnName, 
									@iSourceItemType = ASRSysColumns.dataType
								FROM ASRSysColumns
								INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
								WHERE ASRSysColumns.columnID = @iDBColumnID
					
								IF (@iDBRecord = 0) OR (@iDBRecord = 4) SET @iRecordID = @iInitiatorID
			
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
					'


	SET @sSPCode_1 = '					FROM ASRSysWorkflowInstanceValues IV
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
										FROM ASRSysWorkflowInstanceValues IV
										INNER JOIN ASRSysWorkflowElements Es ON IV.elementID = Es.ID
											AND IV.identifier = Es.identifier
											AND Es.workflowID = @iWorkflowID
											AND Es.identifier = @sRecSelWebFormIdentifier
										WHERE IV.instanceID = @piInstanceID
									END
								END		
			
								IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
								BEGIN
									SET @fValidRecordID = 1

									EXEC spASRWorkflowValidRecord
										@piInstanceID,
										@iDBRecord,
										@iRecordID,
										@sRecSelWebFormIdentifier,
										@sRecSelIdentifier,
										@fValidRecordID	OUTPUT

									IF @fValidRecordID  = 0
									BEGIN
										SET @psMessage = ''''
										SET @pfOK = 0

										RETURN
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
									WHERE ASRSysColumn'


	SET @sSPCode_2 = 's.columnID = @iColumnID
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
										@sCaption + '' - '' + CHAR(13) + 
										''<'' + @sURL + ''?'' + @sQueryString + ''>''
								END
								
								FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption
							END
							CLOSE elementCursor
					
							DEALLOCATE elementCursor
				'


	SET @sSPCode_3 = '		END
					
						DROP TABLE #succeedingElements
					END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2
		+ @sSPCode_3)

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

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetWorkflowQueryString
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

	EXECUTE (@sSPCode_0)

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

	SET @sSPCode_0 = 'Alter PROCEDURE [dbo].spASRGetWorkflowItemValues
			(
				@piElementItemID	integer
			)
			AS
			BEGIN
				DECLARE 
					@iItemType			integer,
					@iLookupColumnID	integer,
					@sDefaultValue		varchar(8000),
					@sTableName			sysname,
					@sColumnName		sysname,
					@iDataType			integer,
					@sSelectSQL			varchar(8000)

				SELECT 			
					@iItemType = ASRSysWorkflowElementItems.itemType,
					@sDefaultValue = ASRSysWorkflowElementItems.inputDefault,
					@iLookupColumnID = ASRSysWorkflowElementItems.lookupColumnID
				FROM ASRSysWorkflowElementItems
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID

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
					CREATE TABLE #dropdownValues 
						([value] varchar(255))

					INSERT INTO #dropdownValues ([value])
						VALUES (null)

					INSERT INTO #dropdownValues ([value])
						SELECT ASRSysWorkflowElementItemValues.value
						FROM ASRSysWorkflowElementItemValues
						WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
						ORDER BY [sequence]

					SELECT [value],
						CASE
							WHEN [value] = @sDefaultValue THEN 1
							ELSE 0
						END
					FROM #dropdownValues 

					DROP TABLE #dropdownValues 
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

					SET @sSelectSQL = ''SELECT null AS [value], 0 UNION SELECT DISTINCT '' + @sTableName + ''.'' + @sColumnName + '' AS [value],''

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
									THEN '''''''' + @sDefaultValue + ''''''''
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

	EXECUTE (@sSPCode_0)

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

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRInstantiateWorkflow
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
			INNER JOIN ASRSysWorkflowElement'


	SET @sSPCode_1 = 's on ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
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

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

	----------------------------------------------------------------------
	-- spASRGetWorkflowFormItems
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'dbo.spASRGetWorkflowFormItems')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRGetWorkflowFormItems]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].spASRGetWorkflowFormItems
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRGetWorkflowFormItems
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
					@iElementType		integer, 
					@iType integer,
					@fValidRecordID	bit,
					@iColumnID	integer
						
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
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID
				FROM ASRSysWorkflowInstances
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				CREATE TABLE #itemValues (ID integer, value varchar(8000), type integer)	
			
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
			
						SET @iType = @iDBColumnDataType
		
						IF (@iDBRecord = 0) OR (@iDBRecord = 4) SET @iRecordID = @iInitiatorID

						IF @iDBRecord = 1
						BEGIN
					'


	SET @sSPCode_1 = '		-- Identified record.
							SELECT @iElementType = ASRSysWorkflowElements.type
							FROM ASRSysWorkflowElements
							WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
								AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sWFFormIdentifier)))
							IF @iElementType = 2
							BEGIN
								 -- WebForm
								SELECT @iRecordID = convert(integer, ISNULL(IV.value, ''0''))
								FROM ASRSysWorkflowInstanceValues IV
								INNER JOIN ASRSysWorkflowElementItems EI ON IV.identifier = EI.identifier
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
									AND IV.identifier = Es.identifier
									AND Es.workflowID = @iWorkflowID
									AND Es.identifier = @sWFFormIdentifier
								WHERE IV.instanceID = @piInstanceID
							END
						END	
						
						IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
						BEGIN
							SET @fValidRecordID = 1

							EXEC spASRWorkflowValidRecord
								@piInstanceID,
								@iDBRecord,
								@iRecordID,
								@sWFFormIdentifier,
								@sWFValueIdentifier,
								@fValidRecordID	OUTPUT

							IF @fValidRecordID = 0
							BEGIN
								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Web Form item record has been deleted.''
											
								SET @psErrorMessage = ''Error loading web form. Web Form item record has been deleted.''
								RETURN
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
			
					INSERT INTO #itemValues (ID, value, type)
					VALUES (@iID, @sValue, @iType)
			
					FETCH NEXT FROM itemCursor IN'


	SET @sSPCode_2 = 'TO @iID, @iItemType, @iDBColumnID, @iDBRecord, @sWFFormIdentifier, @sWFValueIdentifier	
				END
				CLOSE itemCursor
				DEALLOCATE itemCursor
			
				SELECT thisFormItems.*, 
					#itemValues.value, 
					#itemValues.type AS [sourceItemType]
				FROM ASRSysWorkflowElementItems thisFormItems
				LEFT OUTER JOIN #itemValues ON thisFormItems.ID = #itemValues.ID
				WHERE thisFormItems.elementID = @piElementID
				ORDER BY thisFormItems.ZOrder DESC
				DROP TABLE #itemValues
			END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

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
					@iTriggerTableID	int
						
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

				IF @iDataRecord = 4 -- 4 = Triggered record
				BEGIN
					SET @piRecordID = @iInitiatorID
			
					IF @piDataTableID = @iTriggerTableID
					BEGIN
						SET @sIDColumnName = ''ID''
					END
					ELSE
					BEGIN
						SET @sIDColumnName = ''ID_'' + convert(varchar(8000), @iTriggerTableID)
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
						INNER JOIN ASRSysWorkflowElementItems '


	SET @sSPCode_1 = 'EI ON IV.identifier = EI.identifier
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
							AND IV.identifier = Es.identifier
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
			
				SET @fValidRecordID = 1
				IF (@iDataRecord = 0) OR (@iDataRecord = 1) OR (@iDataRecord = 4)
				BEGIN
					EXEC spASRWorkflowValidRecord
						@piInstanceID,
						@iDataRecord,
						@piRecordID,
						@sRecSelWebFormIdentifier,
						@sRecSelIdentifier,
						@fValidRecordID	OUTPUT

					IF @fValidRecordID = 0
					BEGIN
						-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
						EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Stored Data primary record has been deleted.''
						SET @psSQL = ''''
						RETURN
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
					
					IF @iSecondaryDataRecord = 4 -- 4 = Triggered record
					BEGIN
						SET @iSecondaryRecordID = @iInitiatorID
				
						IF @piDataTableID = @iTriggerTableID
						BEGIN
							SET @sSecondaryIDColumnName = ''ID''
						END
						ELSE
						BEGIN
							SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iTriggerTableID)
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
								AND IV.identifier = Es.identifier
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
			'


	SET @sSPCode_2 = '				SET @sSecondaryIDColumnName = ''ID_'' + convert(varchar(8000), @iTempTableID)
						END
					END

					SET @fValidRecordID = 1
					IF (@iSecondaryDataRecord = 0) OR (@iSecondaryDataRecord = 1) OR (@iSecondaryDataRecord = 4)
					BEGIN
						EXEC spASRWorkflowValidRecord
							@piInstanceID,
							@iSecondaryDataRecord,
							@iSecondaryRecordID,
							@sSecondaryRecSelWebFormIdentifier,
							@sSecondaryRecSelIdentifier,
							@fValidRecordID	OUTPUT

						IF @fValidRecordID = 0
						BEGIN
							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC spASRWorkflowActionFailed @piInstanceID, @piElementID, ''Stored Data secondary record has been deleted.''

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
					FROM #dbValues
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

						IF (@iDBRecord = 0) OR (@iDBRecord = 4)
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
								AND IV.columnID is null
								AND CASE
									WHEN WE.type = 5 THEN -- StoredData
										CASE 
											WHEN IV.identifier = WE.identifier THEN 1
											ELSE 0
										END
									ELSE  -- WebForm
										CASE 
											WHEN IV.identifier = @sWFValueIdentifier THEN 1
											ELSE 0
										END
									END = 1
						END

						SET @fValidRecordID = 1
						IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
						BEGIN
							EXEC spASRWorkflowValidRecord
								@piInstanceID,
								@iDBRecord,
								@iRecordID,
								@sWFFormIdentifier,
								@sWFValueIdentifier,
								@fValidRecordID	OUTPUT

							IF @fValidRecordID = 0
							BEGIN
								-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
								EXEC spASRWorkflowAction'


	SET @sSPCode_3 = 'Failed @piInstanceID, @piElementID, ''Stored Data column database value record has been deleted.''

								SET @psSQL = ''''
								RETURN
							END
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

						INSERT INTO ASRSysWorkflowInstanceValues
							(instanceID, elementID, identifier, columnID, value)
							VALUES (@piInstanceID'


	SET @sSPCode_4 = ', @piElementID, '''', @iColumnID, @sValue)
			
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

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].spASRGetWorkflowGridItems
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
					@iElementType		integer, 
					@fValidRecordID	bit,
					@iElementID	integer,
					@iTriggerTableID	integer
			
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
					@iDfltOrderID = ASRSysTables.defaultOrderID,
					@sBaseTableName = ASRSysTables.tableName
				FROM ASRSysWorkflowElementItems
				INNER JOIN ASRSysTables ON ASRSysWorkflowElementItems.tableID = ASRSysTables.tableID
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iWorkflowID = ASRSysWorkflowInstances.workflowID, 
					@iTriggerTableID = ASRSysWorkflows.baseTable
				FROM ASRSysWorkflowInstances
				INNER JOIN ASRSysWorkflows ON ASRSysWorkflowInstances.workflowID = ASRSysWorkflows.ID
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
					'


	SET @sSPCode_1 = 'END
			
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

						SET @iRecordID = @iInitiatorID
					END

					IF @iDBRecord = 4 -- ie. based on the triggered record
					BEGIN
						SET @sSelectSQL = @sSelectSQL + 
							'' WHERE '' + @sBaseTableName + ''.ID_'' + convert(varchar(100), @iTriggerTableID) + '' = '' + convert(varchar(100), @iInitiatorID)

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
								AND IV.identifier = Es.identifier
								AND Es.workflowID = @iWorkflowID
								AND Es.identifier = @sRecSelWebFormIdentifier
							WHERE IV.instanceID = @piInstanceID
						END
			
						SET @sSelectSQL = @sSelectSQL + 
							'' WHERE '' + @sBaseTableName + ''.ID_'' + convert(varchar(100), @iTempTableID) + '' = '' + convert(varchar(100), @iRecordID)
					END
			
					IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
					BEGIN
						SET @fValidRecordID = 1

						EXEC spASRWorkflowValidRecord
							@piInstanceID,
							@iDBRecord,
							@iRecordID,
							@sRecSelWebFormIdentifier,
							@sRecSelIdentifier,
							@fValidRecordID	OUTPUT

						IF @fValidRecordID  = 0
						BEGIN
							SET @pfOK = 0

							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC spASRWorkflowActionFailed @piInstanceID, @iElementID, ''Web Form record selector item record has been deleted.''
							
							-- Need to return a recordset of some kind.
							SELECT '''' AS ''Error''

				'


	SET @sSPCode_2 = '			RETURN
						END
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

	EXECUTE (@sSPCode_0
		+ @sSPCode_1
		+ @sSPCode_2)

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
				@superCursor		cursor	
		
			DECLARE triggeredWFCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT Q.queueID,
				Q.recordID,
				TL.workflowID
			FROM ASRSysWorkflowQueue Q
			INNER JOIN ASRSysWorkflowTriggeredLinks TL ON Q.linkID = TL.linkID
			INNER JOIN ASRSysWorkflows WF ON TL.workflowID = WF.ID
				AND WF.enabled = 1
			WHERE Q.dateInitiated IS null
				AND datediff(dd,DateDue,getdate()) >= 0
		
			OPEN triggeredWFCursor
			FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID
			WHILE (@@fetch_status = 0) 
			BEGIN
				UPDATE ASRSysWorkflowQueue
				SET dateInitiated = getDate()
				WHERE queueID = @iQueueID
				
				-- Create the Workflow Instance record, and remember the ID. */
				INSERT INTO ASRSysWorkflowInstances (workflowID, initiatorID, status, userName)
				VALUES (@iWorkflowID, @iRecordID, 0, ''<Triggered>'')
								
				SELECT @iInstanceID = MAX(id)
				FROM ASRSysWorkflowInstances
				
				-- Create the Workflow Instance Steps records. 
				-- Set the first steps'' status to be 1 (pending Workflow Engine action). 
				-- Set all subsequent steps'' status to be 0 (on hold). */
				SELECT @iStartElementID = ASRSysWorkflowElements.ID
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.type = 0 -- Start element
					AND ASRSysWorkflowElements.workflowID = @iWorkflowID
		
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
					@iInstanceID, 
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
				WHERE ASRSysWorkflowElements.workflowid = @iWorkflowID
				
				DROP TABLE #succeedingElements
		
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
						OR ASRSysWorkflowElementItems.itemType = 7
						OR ASRSysWorkflowElementItems.itemType = 11
						OR ASRSysWorkflowElementItems.itemType = 13
						OR ASRSysWorkflowElementItems.itemType = 14
						OR ASRSysWorkflowElementItems.itemType = 15
						OR ASRSysWorkflowElementItems.itemType = 0)
				UNION
				SELECT  @iInstanceID, ASRSysWorkflowElements.ID, 
					ASRSysWorkflowElements.identifier
				FROM ASRSysWorkflowElements
				WHERE ASRS'


	SET @sSPCode_1 = 'ysWorkflowElements.workflowID = @iWorkflowID
					AND ASRSysWorkflowElements.type = 5						
				
				FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID
			END
			CLOSE triggeredWFCursor
			DEALLOCATE triggeredWFCursor
		END'

	EXECUTE (@sSPCode_0
		+ @sSPCode_1)

	----------------------------------------------------------------------
	-- spASRWorkflowRebuild
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowRebuild]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowRebuild]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowRebuild]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'ALTER PROCEDURE [dbo].spASRWorkflowRebuild
		AS
		BEGIN	
			-- Refresh all scheduled Workflow items in the queue.
			DECLARE @sTableName 	varchar(255),
				@iTableID			int,
				@sSQL				varchar(8000)
			
			-- Get a cursor of the tables in the database.
			DECLARE curTables CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableName, tableID
				FROM ASRSysTables
			OPEN curTables
		
			DELETE FROM ASRSysWorkflowQueue 
			WHERE dateInitiated IS null 
				AND [Immediate] = 0
		
			-- Loop through the tables in the database.
			FETCH NEXT FROM curTables INTO @sTableName, @iTableID
			WHILE @@fetch_status <> -1
			BEGIN
				/* Get a cursor of the records in the current table.  */
				/* Call the Workflow trigger for that table and record  */
				SET @sSQL = ''
					DECLARE @iCurrentID	int,
						@sSQL		varchar(8000)
					
					IF EXISTS (SELECT * FROM sysobjects
					WHERE id = object_id(''''spASRWorkflowRebuild_'' + LTrim(Str(@iTableID)) + '''''') 
						AND sysstat & 0xf = 4)
					BEGIN
						DECLARE curRecords CURSOR FOR
						SELECT id
						FROM '' + @sTableName + ''
		
						OPEN curRecords
		
						FETCH NEXT FROM curRecords INTO @iCurrentID
						WHILE @@fetch_status <> -1
						BEGIN
							SET @sSQL = ''''EXEC spASRWorkflowRebuild_'' + LTrim(Str(@iTableID)) 
								+ '' '''' + convert(varchar(100), @iCurrentID) + ''''''''
							EXEC (@sSQL)
		
							FETCH NEXT FROM curRecords INTO @iCurrentID
						END
						CLOSE curRecords
						DEALLOCATE curRecords
					END''
		
				 EXEC (@sSQL) 
		
				/* Move onto the next table in the database. */ 
				FETCH NEXT FROM curTables INTO @sTableName, @iTableID
			END
		
			CLOSE curTables
			DEALLOCATE curTables
		END'

	EXECUTE (@sSPCode_0)

/* ------------------------------------------------------------- */
PRINT 'Step 4 of 16 - ASR Contact Details'

	update asrsyssystemsettings set settingvalue = '+44 (0)1582 714820'
	where settingvalue = '01582 714820'

	update asrsyssystemsettings set settingvalue = '+44 (0)1582 714814'
	where settingvalue = '01582 714814'


/* ------------------------------------------------------------- */
PRINT 'Step 5 of 16 - Accord Transfer Stored Procedures'

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordNeedToSendAll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordNeedToSendAll]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordPopulateTransaction]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordPopulateTransaction]

	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRAccordNeedToSendAll] 
	(@iTransferType int, 
	@iRecordID int,
	@bResend bit OUTPUT)
	AS
	BEGIN
		SET NOCOUNT ON

		DECLARE @Status integer

		SELECT TOP 1 @Status = Status FROM ASRSysAccordTransactions
		WHERE HRProRecordID = @iRecordID AND TransferType = @iTransferType
		ORDER BY CreatedDateTime DESC
	
		-- Nothing found
		IF @Status IS NULL SET @bResend = 1
	
		-- Previous transaction failed
		IF @Status IN (20) SET @bResend = 0

		--	Previous transaction went as update - should be new
		IF @Status IN (22, 23, 31) SET @bResend = 1
	
		-- Pending, success, or success with warnings, blocked
		IF @Status IN (1, 10, 11, 21, 30) SET @bResend = 0


	END'
	EXEC sp_executesql @NVarCommand

	GRANT EXEC ON [spASRAccordNeedToSendAll] TO [ASRSysGroup]



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
			WHERE HRProRecordID = @piHRProRecordID AND TransferType = @piTransferType
			ORDER BY CreatedDateTime DESC
		
			IF @iStatus IS NULL OR @iStatus = 20 OR @iStatus = 23
			BEGIN
				SET @piTransactionType = 0
				SET @pbSendAllFields = 1
			END
		END

		-- Are we trying to delete something thats never been sent?
		IF @piTransactionType = 2
		BEGIN
			SELECT TOP 1 @iStatus = Status FROM ASRSysAccordTransactions
			WHERE HRProRecordID = @piHRProRecordID AND TransferType = @piTransferType
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

	GRANT EXEC ON [spASRAccordPopulateTransaction] TO [ASRSysGroup]


	-- Archive obsolete Accord transactions
	SET @NVarCommand = 'UPDATE ASRSysAccordTransactions SET Archived = 0 WHERE Archived IS NULL'
	EXEC sp_executesql @NVarCommand


/* ------------------------------------------------------------- */
PRINT 'Step 6 of 16 - Updating Maternity Entitlement'

  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRMaternityExpectedReturn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRMaternityExpectedReturn]

  EXEC('CREATE PROCEDURE [dbo].[spASRMaternityExpectedReturn] (
			@pdblResult datetime OUTPUT,
			@EWCDate datetime,
			@LeaveStart datetime,
			@BabyBirthDate datetime,
			@Ordinary varchar(8000)
			)
			AS
			BEGIN

				IF LOWER(@Ordinary) = ''ordinary''
					IF DateDiff(d,''04/01/2007'', @EWCDate) >= 0
						SET @pdblResult = Dateadd(ww,39,@LeaveStart)
					ELSE IF DateDiff(d,''04/06/2003'', @EWCDate) >= 0
						SET @pdblResult = Dateadd(ww,26,@LeaveStart)
					ELSE
						IF DateDiff(d,''04/30/2000'', @EWCDate) >= 0
							SET @pdblResult = Dateadd(ww,18,@LeaveStart)
						ELSE
							SET @pdblResult = Dateadd(ww,14,@LeaveStart)
				ELSE
					IF DateDiff(d,''04/06/2003'', @EWCDate) >= 0
						SET @pdblResult = Dateadd(ww,52,@LeaveStart)
					ELSE
						--29 weeks from baby birth date (but return on the monday before!)
						SET @pdblResult = DateAdd(d,203 - datepart(dw,DateAdd(d,-2,@BabyBirthDate)),@BabyBirthDate)

			END')

  GRANT EXEC ON [spASRMaternityExpectedReturn] TO [ASRSysGroup]



/* ------------------------------------------------------------- */
PRINT 'Step 7 of 16 - Renaming button to a new ID RE:Fault 11512' 

UPDATE ASRSysUserSettings
SET SettingValue = '20093'
WHERE (SettingValue = '20079')

/* ------------------------------------------------------------- */
PRINT 'Step 8 of 16 - Performance boosting' 

	-- Accord boost
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysAccordTransactions') AND name = N'IDX_HRProRecordID')
		DROP INDEX ASRSysAccordTransactions.[IDX_HRProRecordID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_HRProRecordID] ON [ASRSysAccordTransactions] ([HRProRecordID])'
	EXEC sp_executesql @NVarCommand

	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysAccordTransactionData') AND name = N'IDX_TransactionID')
		DROP INDEX ASRSysAccordTransactionData.[IDX_TransactionID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_TransactionID] ON [ASRSysAccordTransactionData] ([TransactionID], [FieldID])'
	EXEC sp_executesql @NVarCommand

	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysAccordTransactions') AND name = N'IDX_TransactionID')
		DROP INDEX ASRSysAccordTransactions.[IDX_TransactionID]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_TransactionID] ON [ASRSysAccordTransactions] ([TransactionID])'
	EXEC sp_executesql @NVarCommand


	-- CMG boost
	IF EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysAuditTrail') AND name = N'IDX_CMG')
		DROP INDEX ASRSysAuditTrail.[IDX_CMG]
	SET @NVarCommand = 'CREATE NONCLUSTERED INDEX [IDX_CMG] ON [ASRSysAuditTrail] ([ColumnID], [RecordID])'
	EXEC sp_executesql @NVarCommand

	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysAuditTrail') AND name = N'PK_ASRSysAuditTrail')
	BEGIN
		SET @NVarCommand = 'ALTER TABLE dbo.ASRSysAuditTrail ADD CONSTRAINT
					PK_ASRSysAuditTrail PRIMARY KEY CLUSTERED 
					(id) ON [PRIMARY]'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 9 of 16 - Accord Transfer Types' 

	DECLARE @iTransferType nvarchar(2)

	SELECT @NVarCommand = 'ALTER TABLE ASRSysAccordTransferTypes ALTER COLUMN TransferType NVARCHAR(40)'
	EXEC sp_executesql @NVarCommand

	-- Accomodation
	SET @iTransferType = '11'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Accommodation'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Accomodation'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Benefits
	SET @iTransferType = '12'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Benefits'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Benefit'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Bonuses
	SET @iTransferType = '13'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Bonuses'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Bonus'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Commission
	SET @iTransferType = '14'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Commissions'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Commission'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Holiday Sale
	SET @iTransferType = '15'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Holiday Sale'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Holiday Sale'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Insurances
	SET @iTransferType = '16'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Insurance'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Insurance'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Meals
	SET @iTransferType = '17'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Meals'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Meal'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Overtime
	SET @iTransferType = '18'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Overtime'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Overtime'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Pension
	SET @iTransferType = '19'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Pension'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Pension'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Travel
	SET @iTransferType = '20'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Travel'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Travel'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Vehicle
	SET @iTransferType = '21'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Vehicle'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Vehicle'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Weightings
	SET @iTransferType = '22'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - Weightings'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''LWA'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- User Defined
	SET @iTransferType = '23'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 1'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- User Defined
	SET @iTransferType = '24'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 2'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- User Defined
	SET @iTransferType = '25'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 3'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END
	-- User Defined
	SET @iTransferType = '26'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 4'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END
	-- User Defined
	SET @iTransferType = '27'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 5'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- User Defined
	SET @iTransferType = '28'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 6'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- User Defined
	SET @iTransferType = '29'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 7'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- User Defined
	SET @iTransferType = '30'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 8'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- User Defined
	SET @iTransferType = '31'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 9'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- User Defined
	SET @iTransferType = '32'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Allowance - User Defined 10'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',1,''Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Nominal Cost Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (16,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (17,' + @iTransferType + ',1,''Department Code'',0,0,2,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (18,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (19,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Holiday Buy
	SET @iTransferType = '41'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - Holiday Buy'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1,2,''Holiday Purchase'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '42'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 1'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '43'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 2'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '44'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 3'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '45'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 4'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '46'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 5'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '47'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 6'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '48'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 7'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '49'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 8'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '50'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 9'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END

	SET @iTransferType = '51'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Extra Deduction - User Defined 10'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Start Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Deduction Amount'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Reference'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Nominal Amount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Cost Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,' + @iTransferType + ',0,''Cost Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,' + @iTransferType + ',0,''Cost Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,' + @iTransferType + ',0,''Cost Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,' + @iTransferType + ',0,''Cost Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,' + @iTransferType + ',0,''Cost Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,' + @iTransferType + ',0,''Cost Code 7'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,' + @iTransferType + ',0,''Cost Code 8'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,' + @iTransferType + ',0,''Cost Code 9'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (17,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (18,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (19,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (20,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END


	-- Pension Type
	SET @iTransferType = '71'
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = @iTransferType
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, StatusColumnID, IsVisible) VALUES ('+ @iTransferType + ', ''Pension'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,' + @iTransferType + ',1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,' + @iTransferType + ',1,''Employee Code'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,' + @iTransferType + ',1,''Pension Scheme No'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,' + @iTransferType + ',1,''Pension Employee'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,' + @iTransferType + ',0,''Pension Employer'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,' + @iTransferType + ',0,''Pension AVC'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,' + @iTransferType + ',0,''Pension Joining Date'',0,0,2,1,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,' + @iTransferType + ',0,''Pension Leaving Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,' + @iTransferType + ',0,''Pension Policy No'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsEmployeeName) VALUES (9,' + @iTransferType + ',1,''Employee Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentCode) VALUES (10,' + @iTransferType + ',1,''Department Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsDepartmentName) VALUES (11,' + @iTransferType + ',1,''Department Name'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer,IsPayrollCode) VALUES (12,' + @iTransferType + ',1,''Payroll Code'',0,0,2,0,1,1)'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 10 of 16 - Updating Process Checking Procedures'


  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetCurrentUsers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRGetCurrentUsers]

  EXEC('CREATE PROCEDURE [spASRGetCurrentUsers]
    AS
    BEGIN

		SET NOCOUNT ON

		DECLARE @Login nvarchar(200)
		DECLARE @hResult int
		DECLARE @objectToken int
		DECLARE @sTableName nvarchar(50)
		DECLARE @sSQL nvarchar(500)
		DECLARE @bOK bit
		DECLARE @sSQLVersion char(2)

		SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)

		IF @sSQLVersion = ''9''
		BEGIN

			SELECT @Login = [ParameterValue] FROM ASRSysModuleSetup WHERE [ModuleKey] = ''MODULE_SQL''
				AND [ParameterKey] = ''Param_FieldsLoginDetails''

			EXEC @hResult = sp_OACreate ''vbpHRProServer.clsSQLFunctions'', @objectToken OUTPUT
			IF @hResult = 0
			BEGIN

				EXEC sp_ASRUniqueObjectName @sTableName OUTPUT, ''tmp'', 3
				EXEC @hResult = sp_OAMethod @objectToken, ''GetCurrentUsers'', @bOK OUTPUT, @Login, @sTableName

				IF EXISTS (select Name from dbo.sysobjects where id = object_id(@sTableName) and OBJECTPROPERTY(id, N''IsUserTable'') = 1)
				BEGIN
					SET @sSQL = ''SELECT * FROM '' + @sTableName
					EXECUTE sp_executeSQL @sSQL
				END
				ELSE
				BEGIN
					SELECT p.hostname
						   , p.loginame
						   , p.program_name
						   , p.hostprocess
						   , p.sid
						   , p.login_time
						   , p.spid
					FROM     master..sysprocesses p
					JOIN     master..sysdatabases d
					  ON     d.dbid = p.dbid
					WHERE    p.program_name LIKE ''HR Pro%''
					  AND    p.program_name NOT LIKE ''HR Pro Workflow%''
					  AND    d.name = db_name()
					ORDER BY loginame
				END

				EXEC sp_ASRDropUniqueObject @sTableName, 3

			END

			EXEC @hResult = sp_OADestroy @objectToken

		END
		ELSE
		BEGIN

			IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
			BEGIN
				EXEC sp_ASRIntCheckPolls
			END

			SELECT DISTINCT
					 p.hostname
				   , p.loginame
				   , p.program_name
				   , p.hostprocess
				   , p.sid
				   , p.login_time
				   , p.spid
			FROM     master..sysprocesses p
			JOIN     master..sysdatabases d
			  ON     d.dbid = p.dbid
			WHERE    p.program_name LIKE ''HR Pro%''
			  AND    p.program_name NOT LIKE ''HR Pro Workflow%''
			  AND    d.name = db_name()
			ORDER BY loginame
		END

    END')

  GRANT EXEC ON [spASRGetCurrentUsers] TO [ASRSysGroup]


  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetCurrentUsersAppName]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRGetCurrentUsersAppName]

  EXEC('CREATE PROCEDURE spASRGetCurrentUsersAppName
			(
				@psAppName		varchar(8000) OUTPUT,
				@psUserName		varchar(8000)
			)
    AS
    BEGIN

        IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
        BEGIN
            EXEC sp_ASRIntCheckPolls
        END


        SELECT TOP 1
                 @psAppName = rtrim(p.program_name)
        FROM     master..sysprocesses p
        WHERE    p.program_name LIKE ''HR Pro%''
          AND    p.program_name NOT LIKE ''HR Pro Workflow%''
          AND    p.loginame = @psUsername
        GROUP BY p.hostname
               , p.loginame
               , p.program_name
               , p.hostprocess
        ORDER BY loginame

    END')

  GRANT EXEC ON [spASRGetCurrentUsersAppName] TO [ASRSysGroup]


  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetCurrentUsersCountInApp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRGetCurrentUsersCountInApp]

  EXEC('CREATE PROCEDURE [dbo].[spASRGetCurrentUsersCountInApp]
			(
				@piCount		integer		OUTPUT
			)
    AS
    BEGIN

		SET NOCOUNT ON

		DECLARE @Login nvarchar(200)
		DECLARE @hResult int
		DECLARE @objectToken int
		DECLARE @sTableName nvarchar(50)
		DECLARE @sSQL nvarchar(500)
		DECLARE @bOK bit
		DECLARE @sSQLVersion char(2)
		DECLARE @sAppName nvarchar(100)

		SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)
		SET @sAppName = APP_NAME()

		IF @sSQLVersion = ''9''
		BEGIN

			SELECT @Login = [ParameterValue] FROM ASRSysModuleSetup WHERE [ModuleKey] = ''MODULE_SQL''
				AND [ParameterKey] = ''Param_FieldsLoginDetails''

			EXEC @hResult = sp_OACreate ''vbpHRProServer.clsSQLFunctions'', @objectToken OUTPUT
			IF @hResult = 0
			BEGIN
				EXEC @hResult = sp_OAMethod @objectToken, ''CountCurrentUsersInApp'', @piCount OUTPUT, @Login, @sAppName
			END

			EXEC @hResult = sp_OADestroy @objectToken

		END
		ELSE
		BEGIN
			IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
			BEGIN
				EXEC sp_ASRIntCheckPolls
			END

			SELECT @piCount = COUNT(p.Program_Name)
			FROM     master..sysprocesses p
			JOIN     master..sysdatabases d
			  ON     d.dbid = p.dbid
			WHERE    p.program_name = APP_NAME()
			  AND    d.name = db_name()
			GROUP BY p.program_name
		END

    END')

  GRANT EXEC ON [spASRGetCurrentUsersCountInApp] TO [ASRSysGroup]


  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetCurrentUsersCountOnServer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRGetCurrentUsersCountOnServer]

  EXEC('CREATE PROCEDURE spASRGetCurrentUsersCountOnServer
			(
				@piCount		integer OUTPUT,
				@psUserName		varchar(8000)
			)
    AS
    BEGIN

        IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id(''sp_ASRIntCheckPolls'') AND sysstat & 0xf = 4)
        BEGIN
            EXEC sp_ASRIntCheckPolls
        END

        SELECT   @piCount = COUNT(*)
        FROM     master..sysprocesses p
        WHERE    p.program_name LIKE ''HR Pro%''
          AND    p.program_name NOT LIKE ''HR Pro Workflow%''
          AND    p.loginame = @psUsername
        GROUP BY p.hostname
               , p.loginame
               , p.program_name
               , p.hostprocess
        ORDER BY loginame

    END')

  GRANT EXEC ON [spASRGetCurrentUsersCountOnServer] TO [ASRSysGroup]

  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRTestProcessAccount]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRTestProcessAccount]

  EXEC('CREATE PROCEDURE spASRTestProcessAccount
			(
				@psLogin		varchar(4000),
				@pbValid		bit		OUTPUT
			)
    AS
    BEGIN

		SET NOCOUNT ON

		DECLARE @Login nvarchar(200)
		DECLARE @hResult int
		DECLARE @objectToken int
		DECLARE @sTableName nvarchar(50)
		DECLARE @sSQL nvarchar(500)
		DECLARE @bOK bit
		DECLARE @sSQLVersion char(2)

		SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)

		IF @sSQLVersion = ''9''
		BEGIN

			EXEC @hResult = sp_OACreate ''vbpHRProServer.clsSQLFunctions'', @objectToken OUTPUT
			IF @hResult = 0
			BEGIN
				EXEC @hResult = sp_OAMethod @objectToken, ''IsProcessValid'', @pbValid OUTPUT, @psLogin
			END

			EXEC @hResult = sp_OADestroy @objectToken

		END
		ELSE SET @pbValid = 1

    END')

	GRANT EXEC ON [spASRTestProcessAccount] TO [ASRSysGroup]


  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRLockCheck]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[sp_ASRLockCheck]

  EXEC('CREATE PROCEDURE sp_ASRLockCheck AS
	BEGIN

		SET NOCOUNT ON

		DECLARE @sSQLVersion char(2)

		SELECT @sSQLVersion = substring(@@version,charindex(''-'',@@version)+2,1)

		IF @sSQLVersion = ''9'' AND APP_NAME() <> ''HR Pro Workflow Service''
		BEGIN

			CREATE TABLE #tmpSysProcess1 (hostname nvarchar(50), loginname nvarchar(50), program_name nvarchar(50), hostprocess int, sid binary(86), login_time datetime, spid smallint)
			INSERT #tmpSysProcess1 EXEC spASRGetCurrentUsers
		
			SELECT ASRSysLock.* FROM ASRSysLock
			LEFT OUTER JOIN #tmpSysProcess1 syspro 
				ON ASRSysLock.spid = syspro.spid AND ASRSysLock.login_time = syspro.login_time
			WHERE priority = 2 OR syspro.spid IS NOT NULL
			ORDER BY priority

			DROP TABLE #tmpSysProcess1

		END
		ELSE
		BEGIN

			SELECT ASRSysLock.* FROM ASRSysLock
			LEFT OUTER JOIN master..sysprocesses syspro 
				ON asrsyslock.spid = syspro.spid AND asrsyslock.login_time = syspro.login_time
			WHERE Priority = 2 OR syspro.spid IS NOT NULL
			ORDER BY Priority

		END

		SET NOCOUNT OFF

	END')


	GRANT EXEC ON [sp_ASRLockCheck] TO [ASRSysGroup]


/* ------------------------------------------------------------- */
PRINT 'Step 11 of 16 - Removing obsolete procedure'

  --Agreed with TM to remove this Intranet Procedure in the main update script
  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ASRIntCheckUserSessions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[sp_ASRIntCheckUserSessions]


/* ------------------------------------------------------------- */
PRINT 'Step 12 of 16 - sp_OACreate/sp_OADestroy cleanup'

  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetServerDLLVersion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRGetServerDLLVersion]

  EXEC('CREATE PROCEDURE [dbo].[spASRGetServerDLLVersion]
			(
			@strVersion varchar(255) OUTPUT
			)
		AS
		BEGIN

			SET NOCOUNT ON

			DECLARE @objectToken int
			DECLARE @hResult int

			  -- Create Server DLL object
			EXEC @hResult = sp_OACreate ''vbpHRProServer.clsGeneral'', @objectToken OUTPUT
			IF @hResult = 0
				EXEC @hResult = sp_OAMethod @objectToken, ''GetVersion'', @strVersion OUTPUT
			ELSE
				SET @strVersion = ''0.0.0''
				
			EXEC sp_OADestroy @objectToken
			
		END')

	GRANT EXEC ON [spASRGetServerDLLVersion] TO [ASRSysGroup]


  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetDomainPolicy]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRGetDomainPolicy]

  EXEC('CREATE PROCEDURE [dbo].[spASRGetDomainPolicy]
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

			-- Initialise the variables
			SET @LockoutDuration = 0
			SET @lockoutThreshold  = 0
			SET @lockoutObservationWindow  = 0
			SET @maxPwdAge  = 0
			SET @minPwdAge  = 0
			SET @minPwdLength  = 0
			SET @pwdHistoryLength  = 0 
			SET @pwdProperties  = 0

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
		END')

	GRANT EXEC ON [spASRGetDomainPolicy] TO [ASRSysGroup]


  if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRGetDomains]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
  drop procedure [dbo].[spASRGetDomains]

  EXEC('CREATE PROCEDURE [dbo].spASRGetDomains
		(@DomainString varchar(8000) OUTPUT)
	AS
	BEGIN

		SET NOCOUNT ON
	
		DECLARE @objectToken int
		DECLARE @hResult int
		DECLARE @hResult2 varchar(255)
		DECLARE @pserrormessage varchar(255)
	
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
	
	END')


	GRANT EXEC ON [spASRGetDomains] TO [ASRSysGroup]


/* ------------------------------------------------------------- */
PRINT 'Step 13 of 16 - Modifying Workflow system permission descriptions'

	UPDATE ASRSysPermissionItems
	SET description = 'Initiate'
	WHERE itemID = 150

/* ------------------------------------------------------------- */
PRINT 'Step 14 of 16 - Modifying Workflow stored procedures'

	----------------------------------------------------------------------
	-- spASRWorkflowOutOfOfficeConfigured
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRWorkflowOutOfOfficeConfigured]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRWorkflowOutOfOfficeConfigured]

	SET @sSPCode_0 = 'CREATE PROCEDURE [dbo].[spASRWorkflowOutOfOfficeConfigured]
		AS
		BEGIN
			DECLARE @iDummy Int
		END'
	EXECUTE (@sSPCode_0)

	SET @sSPCode_0 = 'Alter PROCEDURE dbo.spASRWorkflowOutOfOfficeConfigured
		(
		    @pfOutOfOfficeConfigured bit output
		)
		AS
		BEGIN
			DECLARE
				@iCount	integer
		
			-- Check if the SP that checks if the current user is OutOfOffice exists
			SELECT @iCount = COUNT(*)
			FROM sysobjects
			WHERE id = object_id(''spASRWorkflowOutOfOfficeCheck'')
				AND sysstat & 0xf = 4
		
			IF @iCount > 0 
			BEGIN
				-- Check if the SP that sets/resets the current user to be OutOfOffice exists
				SELECT @iCount = COUNT(*)
				FROM sysobjects
				WHERE id = object_id(''spASRWorkflowOutOfOfficeSet'')
					AND sysstat & 0xf = 4
			END

			IF @iCount > 0 
			BEGIN
				-- Check if the the Activation column has been defined
				SELECT @iCount = convert(integer, isnull(parameterValue, ''0''))
				FROM ASRSysModuleSetup
				WHERE moduleKey = ''MODULE_WORKFLOW''
					AND parameterKey = ''Param_DelegationActivatedColumn''
			END
		
			SET @pfOutOfOfficeConfigured = 
			CASE	
				WHEN @iCount > 0 THEN 1
				ELSE 0
			END
		END'

	EXECUTE (@sSPCode_0)


/* ------------------------------------------------------------- */
PRINT 'Step 15 of 16 - Accord routines'

	-- 'Resend' security option
	DELETE FROM ASRSysPermissionItems WHERE itemid in (153)
	INSERT INTO ASRSysPermissionItems (ItemID,Description,listOrder,categoryID,itemKey)
		VALUES (153,'Resend Transfer',60,41,'RESEND')

	SELECT @iRecCount = count(*)
	FROM ASRSysGroupPermissions
	WHERE itemid IN (153)

	IF @iRecCount = 0 
	BEGIN
		INSERT ASRSysGroupPermissions (itemid, groupName, permitted)
			SELECT DISTINCT 153, groupName, 1 FROM ASRSysGroupPermissions WHERE ((itemid = 146 AND permitted = 1))
	END

	-- Stored procedures
	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordPopulateTransaction]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordPopulateTransaction]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordVoidPreviousTransactions]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordVoidPreviousTransactions]

	if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spASRAccordSetLatestToType]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
	drop procedure [dbo].[spASRAccordSetLatestToType]


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
			WHERE HRProRecordID = @piHRProRecordID AND TransferType = @piTransferType
			ORDER BY CreatedDateTime DESC
		
			IF @iStatus IS NULL OR @iStatus = 20 OR @iStatus = 23 OR @iStatus = 31
			BEGIN
				SET @piTransactionType = 0
				SET @pbSendAllFields = 1
			END
		END

		-- Are we trying to delete something thats never been sent?
		IF @piTransactionType = 2
		BEGIN
			SELECT TOP 1 @iStatus = Status FROM ASRSysAccordTransactions
			WHERE HRProRecordID = @piHRProRecordID AND TransferType = @piTransferType
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

		SET NOCOUNT OFF

	END
	END'
	EXEC sp_executesql @NVarCommand


	-- spASRAccordVoidPreviousTransactions
	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRAccordVoidPreviousTransactions] (
		@piTransferType int ,
		@piHRProRecordID int)
	AS
	BEGIN	

		SET NOCOUNT ON

		UPDATE ASRSysAccordTransactions SET Status = 31
			WHERE HRProRecordID = @piHRProRecordID AND TransferType = @piTransferType

		SET NOCOUNT OFF

	END'
	EXEC sp_executesql @NVarCommand


	SELECT @NVarCommand = 'CREATE PROCEDURE [dbo].[spASRAccordSetLatestToType] (
		@piTransferType int ,
		@piHRProRecordID int,
		@piTransactionType int)
	AS
	BEGIN	

		SET NOCOUNT ON

		UPDATE ASRSysAccordTransactions SET TransactionType = @piTransactionType
		WHERE TransactionID = (SELECT TOP 1 TransactionID FROM ASRSysAccordTransactions
			WHERE HRProRecordID = @piHRProRecordID AND TransferType = @piTransferType
			ORDER BY CreatedDateTime DESC)

		SET NOCOUNT OFF

	END'
	EXEC sp_executesql @NVarCommand




/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Step 16 of 16 - Updating Versions'

delete from asrsyssystemsettings
where [Section] = 'database' and [SettingKey] = 'version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('database', 'version', '3.3')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '3.3.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.3.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v3.3')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v3.3 Of HR Pro'
