/* --------------------------------------------------- */
/* Update the database from version 8.2 to version 8.3 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(MAX),
	@iSQLVersion numeric(3,1),
	@NVarCommand nvarchar(MAX),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16),
	@sTableName	sysname,
	@sIndexName	sysname,
	@fPrimaryKey	bit;
	
DECLARE @sSPCode nvarchar(MAX)


/* ----------------------------------- */
/* Avoid the (1 Row Affected) messages */
/* ----------------------------------- */
SET NOCOUNT ON;
SET @DBName = DB_NAME();

/* ------------------------------------------------------- */
/* Get the database version from the ASRSysSettings table. */
/* ------------------------------------------------------- */

SELECT @sDBVersion = [SettingValue] FROM ASRSysSystemSettings
where [Section] = 'database' and [SettingKey] = 'version'

/* Exit if the database is not previous or current version . */
/* NB. We allow the script to run even if the database is the new version, as the flags set at the end of the script */
/* may need to be run if we issue corrected versions of the applications without updating the database verion number. */
IF (@sDBVersion <> '8.2') and (@sDBVersion <> '8.3')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2008 or above
SELECT @iSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY('ProductVersion')));
IF (@iSQLVersion < 10)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of OpenHR', 16, 1)
	RETURN
END


/* ------------------------------------------------------- */
PRINT 'Step - Workspace Integration'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetWorkflowIDFromName]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetWorkflowIDFromName];
	EXEC sp_executesql N'CREATE PROCEDURE spASRGetWorkflowIDFromName(
		@name varchar(255),
		@id integer OUTPUT)
	AS
	BEGIN

		IF (SELECT COUNT(id) FROM ASRSysWorkflows WHERE Name = @name) = 1
			SELECT @id = id FROM ASRSysWorkflows WHERE Name = @name;
		ELSE
			SET @id = 0;

	END'


/* ------------------------------------------------------- */
PRINT 'Step - Calculation Updates'
/* ------------------------------------------------------- */

   UPDATE tbstat_componentcode SET [precode] = 'POWER(', [aftercode] = ')', [code] = ', ' WHERE ID = 15 AND isoperator = 1;


/* ------------------------------------------------------- */
PRINT 'Step - Organisation Reports'
/* ------------------------------------------------------- */

	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[ASRSysOrganisationReport]') AND xtype in (N'U'))
	BEGIN

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysOrganisationReport](
			[ID] [int] IDENTITY(1,1) NOT NULL,
			[Name] [varchar](50) NOT NULL,
			[Description] [varchar](255) NOT NULL,
			[BaseViewID] [int] NOT NULL,
			[UserName] [varchar](50) NOT NULL,
			[Timestamp] [timestamp] NOT NULL)';

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysOrganisationReportAccess](
			[GroupName] varchar(256) NOT NULL,
			[Access] varchar(2) NOT NULL,
			[ID] int NOT NULL)';

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysOrganisationColumns](
			[ID] [int] IDENTITY(1,1) NOT NULL,
			[OrganisationID] [int] NOT NULL,
			[ColumnID] [int] NOT NULL,
			[Prefix] [varchar](50) NULL,
			[Suffix] [varchar](50) NULL,
			[FontSize] int,
			[Decimals] int,
			[Height] int,
			[ConcatenateWithNext] bit)';

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysOrganisationReportFilters](
			[ID] [int] IDENTITY(1,1) NOT NULL,
			[OrganisationID] int NOT NULL,
			[FieldID] int NOT NULL,
			[Operator] [int] NOT NULL,
			[Value] nvarchar(MAX) NOT NULL)';

	END


	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysOrganisationColumns', 'U') AND name = 'ViewID')
		EXEC sp_executesql N'ALTER TABLE ASRSysOrganisationColumns ADD ViewID int NULL;';



/* ------------------------------------------------------- */
PRINT 'Step - SQL 2016 Support'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfASRSQLVersion]') AND sysstat & 0xf = 0)
		DROP FUNCTION [dbo].[udfASRSQLVersion]

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfASRSQLVersion]()
	RETURNS integer
	AS
	BEGIN
		RETURN convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY(''ProductVersion'')))
	END'



	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetActualUserDetails]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetActualUserDetails];

	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRGetActualUserDetails]
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

	   SET @sSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY(''ProductVersion'')));

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

	END';


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRInstantiateWorkflow]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRInstantiateWorkflow];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRInstantiateWorkflow]
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
				@sTargetName			nvarchar(MAX) = '''',
				@sSQL					nvarchar(MAX),
				@hResult				integer,
				@sActualLoginName		sysname,
				@fUsesInitiator			bit, 
				@bUseAsTargetIdentifier bit,
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
				@fHasTargetIdentifier bit,
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
				@fSaveForLater			bit,
				@fResult	bit;
		
   	   SET @iSQLVersion = dbo.udfASRSQLVersion();

			DECLARE @succeedingElements table(elementID int);
			DECLARE	@outputTable table (id int NOT NULL);
		
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
				@fEnabled = [enabled],
				@fHasTargetIdentifier = [HasTargetIdentifier]
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
						@iRecordCount OUTPUT,
						@sTargetName OUTPUT;
				END

				IF @fHasTargetIdentifier = 1
					SET @sTargetName = ''<Unidentified>'';
			
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
				[TargetName],
				[parent1TableID],
				[parent1RecordID],
				[parent2TableID],
				[parent2RecordID],
				[pageno])
			OUTPUT inserted.ID INTO @outputTable
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@sTargetName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID,
				0);
						
			SELECT @piInstanceID = id FROM @outputTable;
		
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
							@iStoredDataRecordID	OUTPUT,
							@bUseAsTargetIdentifier OUTPUT,
							@fResult OUTPUT;
		
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
							@fSaveForLater OUTPUT,
							0;
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
		
		END'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRMobileInstantiateWorkflow]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRMobileInstantiateWorkflow];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRMobileInstantiateWorkflow]
		(
			@piWorkflowID	integer,			
			@psKeyParameter	varchar(max),			
			@psPWDParameter	varchar(max),			
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
				@bUseAsTargetIdentifier bit,
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
				@fSaveForLater			bit,
				@fResult	bit;
			
         SELECT @iSQLVersion = dbo.udfASRSQLVersion();

			DECLARE @succeedingElements table(elementID int);
			DECLARE	@outputTable table (id int NOT NULL);
		
			SET @iInitiatorID = 0;
			SET @psFormElements = '''';
			SET @psMessage = '''';
			SET @iParent1TableID = 0;
			SET @iParent1RecordID = 0;
			SET @iParent2TableID = 0;
			SET @iParent2RecordID = 0;
		
			SELECT
			-- @fExternallyInitiated = CASE
			--		WHEN initiationType = 2 THEN 1
			--		ELSE 0
			--	END,
				@fEnabled = enabled
			FROM ASRSysWorkflows
			WHERE ID = @piWorkflowID;

			--IF @fExternallyInitiated = 1
			--BEGIN
				IF @fEnabled = 0
				BEGIN
					/* Workflow is disabled. */
					SET @psMessage = ''This link is currently disabled.'';
					RETURN
				END
		
				SET @sActualLoginName = @psKeyParameter;
			--END
			--ELSE
			--BEGIN
				--SET @sActualLoginName = SUSER_SNAME();
				
				SET @sSQL = ''spASRSysMobileGetCurrentUserRecordID'';
				IF EXISTS (SELECT * FROM sysobjects WHERE type = ''P'' AND name = @sSQL)
				BEGIN
					SET @hResult = 0;
			
					EXEC @hResult = @sSQL 
						@psKeyParameter,			
						@iRecordID OUTPUT,
						@iRecordCount OUTPUT;
				END
			
			print @iRecordID;
			
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
			--END
		
			/* Create the Workflow Instance record, and remember the ID. */
			INSERT INTO [dbo].[ASRSysWorkflowInstances] (workflowID, 
				[initiatorID], 
				[status], 
				[userName], 
				[parent1TableID],
				[parent1RecordID],
				[parent2TableID],
				[parent2RecordID],
				pageno)
			OUTPUT inserted.ID INTO @outputTable
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID,
				0);
						
			SELECT @piInstanceID = id FROM @outputTable;
		
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
							@iStoredDataRecordID	OUTPUT,
							@bUseAsTargetIdentifier OUTPUT,
							@fResult	OUTPUT;
		
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
							@fSaveForLater OUTPUT,
							0;
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
		
		END'


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
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
		@bUseAsTargetIdentifier	bit,
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
					
   SELECT @iSQLVersion = dbo.udfASRSQLVersion();
					
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
		ISNULL(E.trueFlowType, 0),
		ISNULL(E.trueFlowExprID, 0)
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
					@iStoredDataRecordID	OUTPUT,
					@bUseAsTargetIdentifier OUTPUT,
					@fResult OUTPUT;

				IF @fResult = 1
				BEGIN
					IF @iStoredDataAction = 0 -- Insert
					BEGIN
						SET @sSPName  = ''sp_ASRInsertNewRecord''

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
						SET @sSPName  = ''sp_ASRUpdateRecord''

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
									@iStoredDataTableID,
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

						SET @sSPName  = ''sp_ASRDeleteRecord''

						SET @iRetryCount = 0;
						SET @fDeadlock = 1;

						WHILE @fDeadlock = 1
						BEGIN
							SET @fDeadlock = 0;
							SET @iErrorNumber = 0;

							BEGIN TRY
								EXEC @sSPName
									@iResult OUTPUT,
									@iStoredDataTableID,
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

						IF @bUseAsTargetIdentifier = 1
						BEGIN
							EXEC [dbo].[spASRRecordDescription] @iStoredDataTableID, @iStoredDataRecordID, @sEvalRecDesc OUTPUT;
							UPDATE ASRSysWorkflowInstances SET TargetName = @sEvalRecDesc WHERE ID = @piInstanceID;
						END

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
				ELSE				
				BEGIN
					SET @fStoredDataOK = 0;

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
						SET completionCount = isnull(completionCount, 0) - 1
						WHERE instanceID = @piInstanceID
							AND elementID = @iElementID;
					END
					ELSE
					BEGIN
						UPDATE ASRSysWorkflowInstanceSteps
						SET completionCount = isnull(completionCount, 0) - 1
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
				END;
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
END'




PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '8.3';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '8.3.1';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '8.3.1';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'outlook service 2', 'minimum version', '1.0.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v8.3')


/* -------------------------------------------- */
/* Set Refresh flag ? Comment out if not needed */
/* -------------------------------------------- */
EXEC dbo.spsys_setsystemsetting 'database', 'refreshstoredprocedures', 1;


/* ------------------------------------- */
/* Reapply the (1 Row Affected) messages */
/* ------------------------------------- */
SET NOCOUNT OFF;

/* ------------------ */
/* Display OK Message */
/* ------------------ */
PRINT 'Update Script Has Converted Your HR Pro Database To Use v8.3 Of OpenHR'