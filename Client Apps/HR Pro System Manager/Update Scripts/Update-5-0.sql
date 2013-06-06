/* --------------------------------------------------- */
/* Update the database from version 4.3 to version 5.0 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(max),
	@iSQLVersion int,
	@NVarCommand nvarchar(max),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16),
	@sTableName	sysname,
	@sIndexName	sysname,
	@fPrimaryKey	bit;
	
DECLARE @ownerGUID uniqueidentifier,
	@nextid integer,
	@sSPCode nvarchar(max);

DECLARE @admingroups TABLE(groupname nvarchar(255))


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
IF (@sDBVersion <> '4.3') and (@sDBVersion <> '5.0')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2008 or above
SELECT @iSQLVersion = convert(float,substring(@@version,charindex('-',@@version)+2,2))
IF (@iSQLVersion < 9)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of OpenHR', 16, 1)
	RETURN
END

/* ------------------------------------------------------------- */
PRINT 'Step 1 - System procedures'

	IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spstat_setdefaultsystemsetting]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [dbo].[spstat_setdefaultsystemsetting];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRGetAuditTrail]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRGetAuditTrail];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRAuditTable]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRAuditTable];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRAllTablePermissions]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRAllTablePermissions];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRAllTablePermissionsForGroup]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRAllTablePermissionsForGroup];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRFn_GetCurrentUser]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRFn_GetCurrentUser];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRSendMessage]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRSendMessage];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRDelegateWorkflowEmail]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRDelegateWorkflowEmail];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRDropTempObjects]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRDropTempObjects];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRUpdateStatistics]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRUpdateStatistics];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersAppName]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersAppName];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersCountOnServer]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersCountOnServer];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetDomainPolicy]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetDomainPolicy];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRSubmitWorkflowStep]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRSubmitWorkflowStep];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_audittable]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_audittable];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_getaudittrail]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_getaudittrail];



	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRDropTempObjects]
	AS
	BEGIN

		DECLARE	@sObjectName varchar(255),
				@sUsername varchar(255),
				@sXType varchar(50);
				
		DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
		SELECT [dbo].[sysobjects].[name], [sys].[schemas].[name], [dbo].[sysobjects].[xtype]
		FROM [dbo].[sysobjects] 
				INNER JOIN [sys].[schemas]
				ON [dbo].[sysobjects].[uid] = [sys].[schemas].[schema_id]
		WHERE LOWER([sys].[schemas].[name]) != ''dbo'' AND LOWER([sys].[schemas].[name]) != ''fusion''
				AND (OBJECTPROPERTY(id, N''IsUserTable'') = 1
					OR OBJECTPROPERTY(id, N''IsProcedure'') = 1
					OR OBJECTPROPERTY(id, N''IsInlineFunction'') = 1
					OR OBJECTPROPERTY(id, N''IsScalarFunction'') = 1
					OR OBJECTPROPERTY(id, N''IsTableFunction'') = 1);

		OPEN tempObjects;
		FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType;
		WHILE (@@fetch_status <> -1)
		BEGIN		
			IF UPPER(@sXType) = ''U''
				-- user table
				BEGIN
					EXEC (''DROP TABLE ['' + @sUsername + ''].['' + @sObjectName + '']'');
				END

			IF UPPER(@sXType) = ''P''
				-- procedure
				BEGIN
					EXEC (''DROP PROCEDURE ['' + @sUsername + ''].['' + @sObjectName + '']'');
				END

			IF UPPER(@sXType) = ''TF''
				-- UDF
				BEGIN
					EXEC (''DROP FUNCTION ['' + @sUsername + ''].['' + @sObjectName + '']'');
				END

			IF UPPER(@sXType) = ''FN''
				-- UDF
				BEGIN
					EXEC (''DROP FUNCTION ['' + @sUsername + ''].['' + @sObjectName + '']'');
				END
		
			FETCH NEXT FROM tempObjects INTO @sObjectName, @sUsername, @sXType;
		
		END
		CLOSE tempObjects;
		DEALLOCATE tempObjects;
	
		EXEC (''DELETE FROM [dbo].[ASRSysSQLObjects]'');


		-- Clear out any temporary tables that may have got left behind from the createunique function
		DECLARE tempObjects CURSOR LOCAL FAST_FORWARD FOR 
		SELECT [dbo].[sysobjects].[name]
		FROM [dbo].[sysobjects] 
		INNER JOIN [dbo].[sysusers]	ON [dbo].[sysobjects].[uid] = [dbo].[sysusers].[uid]
		LEFT JOIN ASRSysTables ON sysobjects.[name] = ASRSysTables.TableName
		WHERE LOWER([dbo].[sysusers].[name]) = ''dbo''
			AND OBJECTPROPERTY(sysobjects.id, N''IsUserTable'') = 1
			AND ASRSysTables.TableName IS NULL
			AND [dbo].[sysobjects].[name] LIKE ''tmp%'';

		OPEN tempObjects;
		FETCH NEXT FROM tempObjects INTO @sObjectName;
		WHILE (@@fetch_status <> -1)
		BEGIN		
			EXEC (''DROP TABLE [dbo].['' + @sObjectName + '']'');
			FETCH NEXT FROM tempObjects INTO @sObjectName;
		END

		CLOSE tempObjects;
		DEALLOCATE tempObjects;

	END'


	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRUpdateStatistics]
	AS
	BEGIN

		SET NOCOUNT ON;

		DECLARE @sTableName		nvarchar(255),
				@sSchema		nvarchar(255),
				@sVarCommand	nvarchar(MAX);

		-- Checking fragmentation
		DECLARE tables CURSOR FOR
			SELECT sc.[Name], so.[Name]
			FROM sys.sysobjects so
				INNER JOIN sys.sysindexes si ON so.id = si.id
				INNER JOIN sys.schemas sc ON so.uid  = sc.schema_id
			WHERE so.type =''U'' AND si.indid < 2 AND si.rows > 0
			ORDER BY sc.name, so.[Name];

		-- Open the cursor
		OPEN tables;

		-- Loop through all the tables in the database running dbcc showcontig on each one
		FETCH NEXT FROM tables INTO @sSchema, @sTableName;

		WHILE @@FETCH_STATUS = 0
		BEGIN
			SET @sVarCommand = ''UPDATE STATISTICS ['' + @sSchema + ''].['' + @sTableName + ''] WITH FULLSCAN'';
			EXECUTE sp_executeSQL @sVarCommand;
			FETCH NEXT FROM tables INTO @sSchema, @sTableName;
		END

		-- Close and deallocate the cursor
		CLOSE tables;
		DEALLOCATE tables;

	END'


	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]
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
		
						IF LEN(@sStoredDataSQL) > 0 
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
		END';


	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRAllTablePermissions] 
	(
		@psSQLLogin		varchar(200)
	)
	AS
	BEGIN

		SET NOCOUNT ON;

		/* Return parameters showing what permissions the current user has on all of the tables. */
		DECLARE @iUserGroupID	int;

		/* Initialise local variables. */
		SELECT @iUserGroupID = usg.gid
		FROM sysusers usu
		left outer join
		(sysmembers mem inner join sysusers usg on mem.groupuid = usg.uid) on usu.uid = mem.memberuid
		WHERE (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1) and
			(usg.issqlrole = 1 or usg.uid is null) and
			usu.name = @psSQLLogin AND not (usg.name like ''ASRSys%'');

		-- Cached cut down view of the sysprotects table
		DECLARE @SysProtects TABLE([ID] int, [Action] tinyint, [ProtectType] tinyint, [Columns] varbinary(8000))
		INSERT @SysProtects
			SELECT [ID],[Action],[ProtectType], [Columns] FROM sysprotects
			WHERE [UID] = @iUserGroupID;

		-- Cached version of the Base table IDs
		DECLARE @BaseTableIDs TABLE([ID] int PRIMARY KEY CLUSTERED, [BaseTableID] int)
		INSERT @BaseTableIDs
			SELECT DISTINCT o.ID, v.TableID
			FROM sysobjects o
			INNER JOIN dbo.ASRSysChildViews2 v ON v.ChildViewID = CONVERT(integer,SUBSTRING(o.Name,9,PATINDEX ( ''%#%'' , o.Name) - 9))
			WHERE Name LIKE ''ASRSYSCV%'';


		SELECT o.name, p.action, bt.BaseTableID
		FROM @SysProtects p
		INNER JOIN sysobjects o ON p.id = o.id
		LEFT OUTER JOIN @BaseTableIDs bt ON o.id = bt.id
		WHERE p.protectType <> 206
			AND p.action <> 193
			AND (o.xtype = ''u'' or o.xtype = ''v'')
			AND (o.Name NOT LIKE ''ASRSYS%'' OR o.Name LIKE ''ASRSYSCV%'')
		UNION
		SELECT o.name, 193, bt.BaseTableID
		FROM syscolumns
		INNER JOIN @SysProtects p ON (syscolumns.id = p.id
			AND p.action = 193 
			AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
			AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0)))
		INNER JOIN sysobjects o ON o.id = p.id
		LEFT OUTER JOIN @BaseTableIDs bt ON o.id = bt.id
		WHERE syscolumns.name = ''timestamp''
			AND p.protectType IN (204, 205) 
			AND (o.Name NOT LIKE ''ASRSYS%'' OR o.Name LIKE ''ASRSYSCV%'')
		ORDER BY o.name;

	END'

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRAllTablePermissionsForGroup]
	(
		@psGroupName sysname
	)
	AS
	BEGIN
		-- Return parameters showing what permissions the current user has on all of the tables.
		DECLARE @iUserGroupID	integer;

		-- Initialise local variables.
		SELECT @iUserGroupID = sysusers.gid
		FROM sysusers
		WHERE sysusers.name = @psGroupName;

		SELECT sysobjects.name, sysprotects.action
		FROM sysprotects 
		INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
		WHERE sysprotects.uid = @iUserGroupID
			AND sysprotects.protectType <> 206
			AND (sysobjects.xtype = ''u'' or sysobjects.xtype = ''v'')
			AND (sysobjects.Name NOT LIKE ''ASRSYS%'' OR sysobjects.Name LIKE ''ASRSYSCV%'')
		ORDER BY sysobjects.name;
	
	END'		

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRFn_GetCurrentUser]
	(
		@psResult	varchar(255) OUTPUT
	)
	AS
	BEGIN
		SET @psResult = 
			CASE 
				WHEN UPPER(LEFT(APP_NAME(), 15)) = ''OPENHR WORKFLOW'' THEN ''OpenHR Workflow'' 
				ELSE SUSER_SNAME()
			END;
	END'

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRSendMessage] 
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

		/* Get a cursor of the other logged in users. */
		DECLARE logins_cursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT DISTINCT spid, loginame, uid, login_time
			FROM #tblCurrentUsers
			WHERE (spid <> @@spid and spid <> @Realspid)
			AND (@psSPIDS = '''' OR charindex('' ''+convert(varchar,spid)+'' '', @psSPIDS)>0);

		OPEN logins_cursor;
		FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime;
		WHILE (@@fetch_status = 0)
		BEGIN
			/* Create a message record for each user. */
			INSERT INTO ASRSysMessages 
				(loginname, [message], loginTime, [dbid], [uid], spid, messageTime, messageFrom, messageSource) 
				VALUES(@sLoginName, @psMessage, @dtLoginTime, @iDBid, @iUid, @iSPid, getdate(), @sCurrentUser, @sCurrentApp);

			FETCH NEXT FROM logins_cursor INTO @iSPid, @sLoginName, @iUid, @dtLoginTime;
		END
		CLOSE logins_cursor;
		DEALLOCATE logins_cursor;

		IF OBJECT_ID(''tempdb..#tblCurrentUsers'', N''U'') IS NOT NULL
			DROP TABLE #tblCurrentUsers;

	END'

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spstat_setdefaultsystemsetting](
			@section AS varchar(50),
			@settingkey AS varchar(50),
			@settingvalue AS nvarchar(MAX))
		AS
		BEGIN
			IF NOT EXISTS(SELECT [SettingValue] FROM [asrsyssystemsettings] WHERE [Section] = @section AND [SettingKey] = @settingkey)
				INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue]) VALUES (@section, @settingkey, @settingvalue);	
		END';

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRDelegateWorkflowEmail] 
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
		DECLARE	@sTo				varchar(MAX),
			@sAddress			varchar(MAX),
			@iInstanceID		integer,
			@curRecipients		cursor,
			@sEmailAddress		varchar(MAX),
			@fDelegated			bit,
			@sDelegatedTo		varchar(MAX),
			@fIsDelegate		bit,
			@sTemp		varchar(MAX),
			@fCopyDelegateEmail		bit;

		SET @psMessage = isnull(@psMessage, '''');
		SET @psMessage_HypertextLinks = isnull(@psMessage_HypertextLinks, '''');
		IF (len(ltrim(rtrim(@psTo))) = 0) RETURN;

		-- Get the instanceID of the given step
		SELECT @iInstanceID = instanceID
		FROM dbo.ASRSysWorkflowInstanceSteps
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
		WHERE len(ltrim(rtrim(emailAddress))) > 0
			AND delegated = 0
			AND ltrim(rtrim(emailAddress))  NOT IN
				(SELECT ltrim(rtrim(emailAddress))
				FROM @recipients
				WHERE len(ltrim(rtrim(emailAddress))) > 0
				AND delegated = 1);

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
				''OpenHR Workflow'',
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
				''OpenHR Workflow'',
				1,
				0, 
				@psCopyTo,
				''You have been copied in on the following OpenHR Workflow email with recipients:'' + CHAR(13)
					+ CHAR(9) + @sTo + CHAR(13)	+ CHAR(13)
					+ @psMessage,
				@iInstanceID,
				@psEmailSubject);
		END

		SET @fCopyDelegateEmail = 1
		SELECT @sTemp = LTRIM(RTRIM(UPPER(ISNULL(parameterValue, ''''))))
		FROM ASRSysModuleSetup
		WHERE moduleKey = ''MODULE_WORKFLOW''
			AND parameterKey = ''Param_CopyDelegateEmail''

		IF @sTemp = ''TRUE''
		BEGIN
			DECLARE toCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT ltrim(rtrim(emailAddress)), 
					ltrim(rtrim(delegatedTo))
				FROM @recipients
				WHERE len(ltrim(rtrim(emailAddress))) > 0
				AND delegated = 1;
			
			OPEN toCursor;
			FETCH NEXT FROM toCursor INTO @sAddress, @sDelegatedTo;
			WHILE (@@fetch_status = 0)
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
					''OpenHR Workflow'',
					1,
					0, 
					@sAddress,
					''The following email has been delegated to '' + @sDelegatedTo + char(13) + 
						''--------------------------------------------------'' + char(13) +
						@psMessage + @psMessage_HypertextLinks,
					@iInstanceID,
					@psEmailSubject);

				
				FETCH NEXT FROM toCursor INTO @sAddress, @sDelegatedTo;
			END
			CLOSE toCursor;
			DEALLOCATE toCursor;
		END
	END'

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersAppName]
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
		WHERE p.program_name LIKE ''OpenHR%''
			AND	p.program_name NOT LIKE ''OpenHR Workflow%''
			AND	p.program_name NOT LIKE ''OpenHR Outlook%''
			AND	p.program_name NOT LIKE ''OpenHR Server.Net%''
			AND	p.program_name NOT LIKE ''OpenHR Intranet Embedding%''
			AND	p.loginame = @psUsername
		GROUP BY p.hostname
			   , p.loginame
			   , p.program_name
			   , p.hostprocess
		ORDER BY p.loginame;

	END'

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRGetCurrentUsersCountOnServer]
	(
		@iLoginCount	integer OUTPUT,
		@psLoginName	varchar(MAX)
	)
	AS
	BEGIN

		DECLARE @sSQLVersion	integer,
				@Mode			smallint;

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
			WHERE p.program_name LIKE ''OpenHR%''
				AND	p.program_name NOT LIKE ''OpenHR Workflow%''
				AND	p.program_name NOT LIKE ''OpenHR Outlook%''
				AND	p.program_name NOT LIKE ''OpenHR Server.Net%''
				AND	p.program_name NOT LIKE ''OpenHR Intranet Embedding%''
				AND p.loginame = @psLoginName;
		END
	END';

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRGetDomainPolicy]
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

		SET NOCOUNT ON;

		-- Initialise the variables
		SET @LockoutDuration = 0;
		SET @lockoutThreshold  = 0;
		SET @lockoutObservationWindow  = 0;
		SET @maxPwdAge  = 0;
		SET @minPwdAge  = 0;
		SET @minPwdLength  = 0;
		SET @pwdHistoryLength  = 0;
		SET @pwdProperties  = 0;

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
			, @pwdHistoryLength OUT, @pwdProperties OUT;

	END';

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRSubmitWorkflowStep]
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
			SET @iIndex1 = charindex(CHAR(9), @psFormInput1);
			SET @sValue = LEFT(@psFormInput1, @iIndex1-1);
			SET @sTemp = SUBSTRING(@psFormInput1, @iIndex1+1, LEN(@psFormInput1) - @iIndex1);

			SET @sValueDescription = '''';
			SET @sMessage = ''Successfully '' +
				CASE
					WHEN @iDataAction = 0 THEN ''inserted''
					WHEN @iDataAction = 1 THEN ''updated''
					ELSE ''deleted''
				END + '' record'';

			IF @iDataAction = 2 -- Deleted - Record Description calculated before the record was deleted.
			BEGIN
				SET @sValueDescription = @sTemp;
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
				exec [dbo].[spASREmailImmediate] ''OpenHR Workflow'';
			END
		END
	END'

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spstat_audittable] (
		@piTableID int,
		@piRecordID int,
		@psRecordDesc varchar(255),
		@psValue varchar(MAX))
	AS
	BEGIN	
		DECLARE @sTableName varchar(128);

		-- Get the table name for the given column.
		SELECT @sTableName = tableName 
			FROM dbo.ASRSysTables
			WHERE tableID = @piTableID;

		IF @sTableName IS NULL SET @sTableName = ''<Unknown>'';

		-- Insert a record into the Audit Trail table.
		INSERT INTO dbo.ASRSysAuditTrail 
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
				WHEN UPPER(LEFT(APP_NAME(), 15)) = ''OPENHR WORKFLOW'' THEN ''OpenHR Workflow''
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
	END'

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spstat_getaudittrail] (
		@piAuditType	int,
		@psOrder 		varchar(MAX))
	AS
	BEGIN

		SET NOCOUNT ON;

		DECLARE @sSQL			varchar(MAX),
				@sExecString	nvarchar(MAX);

		IF @piAuditType = 1
		BEGIN

			SET @sSQL = ''SELECT userName AS [User], 
				dateTimeStamp AS [Date / Time], 
				tableName AS [Table], 
				columnName AS [Column], 
				oldValue AS [Old Value], 
				newValue AS [New Value], 
				recordDesc AS [Record Description],
				id
				FROM dbo.ASRSysAuditTrail '';

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
				FROM dbo.ASRSysAuditPermissions '';

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
				FROM dbo.ASRSysAuditGroup '';

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
				HRProModule AS [Module],
				Action AS [Action], 
				id
				FROM dbo.ASRSysAuditAccess '';

			IF LEN(@psOrder) > 0
				SET @sExecString = @sSQL + @psOrder;
			ELSE
				SET @sExecString = @sSQL;

		END

		-- Retreive selected data
		IF LEN(@sExecString) > 0 EXECUTE sp_executeSQL @sExecString;

	END'


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
						IF (len(@sValue) = 0) OR (@sValue = ''null'')
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


/* ------------------------------------------------------------- */
PRINT 'Step 2 - Set default values'

	EXEC dbo.[spstat_setdefaultsystemsetting] 'integration', 'payroll', 'OpenPay';



/* ------------------------------------------------------------- */
PRINT 'Step 3 - System indexes'

	IF  EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysOutlookEvents]') AND name = N'IDX_LinkRecordID')
		DROP INDEX [IDX_LinkRecordID] ON [dbo].[ASRSysOutlookEvents] WITH ( ONLINE = OFF )
	EXEC sp_executesql N'CREATE CLUSTERED INDEX [IDX_LinkRecordID] ON [dbo].[ASRSysOutlookEvents] ([RecordID] ASC, [FolderID] ASC, [TableID] ASC, [LinkID] ASC)'

	IF  EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[tbsys_scriptedobjects]') AND name = N'IDX_TargetObjectID')
		DROP INDEX [IDX_TargetObjectID] ON [dbo].[tbsys_scriptedobjects] WITH ( ONLINE = OFF )
	EXEC sp_executesql N'CREATE NONCLUSTERED INDEX [IDX_TargetObjectID] ON [dbo].[tbsys_scriptedobjects] ([targetid] ASC, [objecttype] ASC) INCLUDE ([lastupdated],	[lastupdatedby], [effectivedate], [locked])'


/* ------------------------------------------------------------- */
PRINT 'Step 4 - Workflow Tab Strips'

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowElementItems', 'U') AND name = 'pageno')
		BEGIN
			EXEC sp_executesql N'ALTER TABLE ASRSysWorkflowElementItems ADD pageno integer NULL;';
			EXEC sp_executesql N'UPDATE ASRSysWorkflowElementItems SET pageno = 0;';
		END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowElementItems', 'U') AND name = 'buttonstyle')
		EXEC sp_executesql N'ALTER TABLE ASRSysWorkflowElementItems ADD buttonstyle tinyint NULL;';

/* ------------------------------------------------------------- */
PRINT 'Step 5 - New Shared Table Transfer Types for NFP'

	-- Pay Scale Group
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 61
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (61, ''Pay Scale Group'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,61,1,''Pay Scale Group'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,61,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,61,1,''Effective Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,61,1,''Increment Type'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,61,0,''Increment Cut Off Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,61,0,''Increment Due Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,61,0,''Increment Period'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,61,0,''Auto Step New Start'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,61,0,''Auto Step'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,61,0,''Payment Level'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,61,0,''Weekly Payslip Display'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,61,0,''Negotiating Body'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,61,0,''Hours per Week'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Pay Scale
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 62
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (62, ''Pay Scale'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,62,1,''Pay Scale Group'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,62,1,''Pay Scale'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,62,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,62,1,''Effective Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,62,1,''Minimum Point'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,62,1,''Maximum Point'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,62,0,''Bar Point'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Point
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 63
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (63, ''Point'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,63,1,''Pay Scale Group'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,63,1,''Pay Scale'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,63,1,''Point'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,63,1,''Effective Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,63,0,''Annual'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,63,0,''Monthly'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,63,0,''Weekly'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,63,0,''Hourly'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Post
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 64
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (64, ''Post'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,64,1,''Post ID'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,64,1,''Post Title'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,64,1,''Effective Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,64,0,''Pay Scale Group'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,64,0,''Pay Scale'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,64,0,''Minimum Point'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,64,0,''Maximum Point'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,64,0,''Bar Point'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,64,0,''Contract Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,64,0,''Full or Part Time'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,64,0,''Post End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,64,0,''In Use'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,64,0,''Cost Centre'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,64,0,''Reports To'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,64,0,''Post Status'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,64,0,''Location'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,64,0,''Duty Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (17,64,0,''Budget FTE'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (18,64,0,''Budget Headcount'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (19,64,0,''Budget Cost'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Appointment
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 65
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (65, ''Appointment'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,65,1,''Post ID'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,65,1,''Company Code'',1,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,65,1,''Staff Number'',0,1,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,65,1,''Effective Date'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,65,1,''Point'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,65,0,''Primary Job'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,65,1,''Protected Group'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,65,1,''Protected Scale'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,65,1,''Protected Point'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,65,1,''Appointment Reason'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,65,1,''Appointment Information'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,65,1,''Auto Increment'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,65,1,''Hours per Week'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,65,1,''Contract Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,65,1,''Full or Part Time'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,65,1,''Appointment End Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,65,1,''Next Review Date'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Negotiating Body
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 66
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (66, ''Negotiating Body'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,66,1,''Code Table ID'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,66,1,''Negotiating Body'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,66,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,66,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,66,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,66,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,66,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,66,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,66,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,66,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,66,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,66,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,66,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,66,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,66,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Post Status
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 67
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (67, ''Post Status'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,67,1,''Code Table ID'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,67,1,''Post Status'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,67,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,67,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Location
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 68
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (68, ''Location'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,68,1,''Code Table ID'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,68,1,''Location'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,68,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,68,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,68,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,68,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,68,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,68,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,68,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,68,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,68,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,68,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,68,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,68,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,68,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Duty Type
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 69
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (69, ''Duty Type'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,69,1,''Code Table ID'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,69,1,''Duty Type'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,69,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,69,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
	END

	-- Appointment Information
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 70
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (70, ''Appointment Information'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,70,1,''Code Table ID'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,70,1,''Appointment Information'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,70,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,70,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
	END


	-- New mappings on employee for NFP
	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 57) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Known As''  WHERE TransferTypeID = 0 AND TransferFieldID = 57'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 58) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Additional Email''  WHERE TransferTypeID = 0 AND TransferFieldID = 58'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 59) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Pension Scheme''  WHERE TransferTypeID = 0 AND TransferFieldID = 59'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 60) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''OMP Scheme''  WHERE TransferTypeID = 0 AND TransferFieldID = 60'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 61) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''P11d''  WHERE TransferTypeID = 0 AND TransferFieldID = 61'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 62) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Personnel No''  WHERE TransferTypeID = 0 AND TransferFieldID = 62'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 63) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Hours Per Day''  WHERE TransferTypeID = 0 AND TransferFieldID = 63'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 64) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Hours Per Month''  WHERE TransferTypeID = 0 AND TransferFieldID = 64'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 65) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Reports To (1)''  WHERE TransferTypeID = 0 AND TransferFieldID = 65'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 66) = 'Unused'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Reports To (2)''  WHERE TransferTypeID = 0 AND TransferFieldID = 66'
		EXEC sp_executesql @NVarCommand
	END

/* ------------------------------------------------------------- */
PRINT 'Step 6 - Changes to Shared Table Transfer for RTI'
	
	-- Update existing columns for Employee transfer
	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 195) = 'Analysis Code 1'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Passport Number''  WHERE TransferTypeID = 0 AND TransferFieldID = 195'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 196) = 'Analysis Code 2'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Work Status''  WHERE TransferTypeID = 0 AND TransferFieldID = 196'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 197) = 'Analysis Code 3'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''EPM6 (Modified) Scheme''  WHERE TransferTypeID = 0 AND TransferFieldID = 197'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 198) = 'Analysis Code 4'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''EEA/Commonwealth Citizen''  WHERE TransferTypeID = 0 AND TransferFieldID = 198'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 199) = 'Analysis Code 5'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Starter Statement A''  WHERE TransferTypeID = 0 AND TransferFieldID = 199'
		EXEC sp_executesql @NVarCommand
	END

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 200) = 'Analysis Code 6'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Starter Statement B''  WHERE TransferTypeID = 0 AND TransferFieldID = 200'
		EXEC sp_executesql @NVarCommand
	END

	-- Add new mappings for Employee transfer
	SELECT @iRecCount = count(TransferFieldID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferFieldID = 201 AND TransferTypeID = 0
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (201,0,0,''Starter Statement C'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (202,0,0,''Irregular Payment Pattern'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (203,0,0,''Student Loan Indicator'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Update existing column for ASPP (Adoption) transfer
	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 77 AND TransferFieldID = 10) = 'Adopter Name'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Adopter Surname'', Mandatory = 1 WHERE TransferTypeID = 77 AND TransferFieldID = 10'
		EXEC sp_executesql @NVarCommand
	END

	-- Add new mappings for ASPP (Adoption) transfer
	SELECT @iRecCount = count(TransferFieldID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferFieldID = 25 AND TransferTypeID = 77
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (25,77,1,''Adopter Forename 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (26,77,0,''Adopter Forename 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END

	-- Update existing column for ASPP (Birth) transfer
	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 78 AND TransferFieldID = 10) = 'Mother Name'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Mother Surname'', Mandatory = 1 WHERE TransferTypeID = 78 AND TransferFieldID = 10'
		EXEC sp_executesql @NVarCommand
	END

	-- Add new mappings for ASPP (Birth) transfer
	SELECT @iRecCount = count(TransferFieldID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferFieldID = 25 AND TransferTypeID = 78
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (25,78,1,''Mother Forename 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (26,78,0,''Mother Forename 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	END


/* ------------------------------------------------------------- */
PRINT 'Step 7 - Modify Workflow Table - Add PictureID Column '

/* ASRSysWorkflowElements - Add new Attachment_DBColumnID column */
SELECT @iRecCount = COUNT(id) FROM syscolumns
WHERE id = OBJECT_ID('tbsys_Workflows', 'U')
AND name = 'PictureID'

IF @iRecCount = 0
BEGIN
	SELECT @NVarCommand = 'ALTER TABLE tbsys_Workflows ADD 
						PictureID [int] NULL'
	EXEC sp_executesql @NVarCommand
END

	EXEC sp_executesql N'UPDATE tbsys_Workflows SET pictureid = NULL WHERE pictureid = 0';


/* ------------------------------------------------------------- */
PRINT 'Step 8 - New Mobile User Logins Table'

	IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysGroups]'))
	BEGIN
		EXEC sp_executesql N'CREATE VIEW [dbo].[ASRSysGroups] AS
			SELECT uid AS ID, name AS Name
			FROM sys.sysusers
			WHERE (gid = uid) AND (gid > 0) AND (NOT (name LIKE ''ASRSys%'')) AND (NOT (name LIKE ''db[_]%''));'
	END
	
	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbsys_mobilelogins]') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[tbsys_mobilelogins](
			[userid] [integer] NOT NULL,
			[password] [nvarchar](max) NULL,
			[newpassword] [nvarchar](max) NULL);';
	END
	
	IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[ASRSysWorkflows]'))
	DROP VIEW [dbo].[ASRSysWorkflows]
	
	EXEC sp_executesql N'CREATE VIEW [dbo].[ASRSysWorkflows]
					WITH SCHEMABINDING
					AS SELECT base.[id], base.[name], base.[description], base.[enabled], base.[initiationtype], base.[basetable], base.[querystring], base.[pictureid], obj.[locked], obj.[lastupdated], obj.[lastupdatedby]
						FROM dbo.[tbsys_workflows] base
						INNER JOIN dbo.[tbsys_scriptedobjects] obj ON obj.targetid = base.id AND obj.objecttype = 10
						INNER JOIN dbo.[tbstat_effectivedates] dt ON dt.[type] = 1
						WHERE obj.effectivedate <= dt.[date]'
						
	EXEC sp_executesql N'CREATE TRIGGER [dbo].[DEL_ASRSysWorkflows] ON [dbo].[ASRSysWorkflows]
					INSTEAD OF DELETE
					AS
					BEGIN
						SET NOCOUNT ON;

						DELETE FROM [tbsys_workflows] WHERE id IN (SELECT id FROM deleted);
					END'

	EXEC sp_executesql N'CREATE TRIGGER [dbo].[INS_ASRSysWorkflows] ON [dbo].[ASRSysWorkflows]
					INSTEAD OF INSERT
					AS
					BEGIN
	
						SET NOCOUNT ON;
	
						-- Update objects table
						IF NOT EXISTS(SELECT [guid]
							FROM dbo.[tbsys_scriptedobjects] o
							INNER JOIN inserted i ON i.id = o.targetid AND o.objecttype = 10)
						BEGIN
							INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
								SELECT NEWID(), 10, [id], dbo.[udfsys_getownerid](), ''01/01/1900'',1,0, GETDATE()
									FROM inserted;
						END

						-- Update base table								
						INSERT dbo.[tbsys_workflows] ([id], [name], [description], [enabled], [initiationType], [baseTable], [queryString], [pictureid]) 
							SELECT [id], [name], [description], [enabled], [initiationType], [baseTable], [queryString], [pictureid] FROM inserted;

					END'

	EXEC sp_executesql N'GRANT SELECT,INSERT,UPDATE,DELETE ON [ASRSysWorkflows] TO [ASRSysGroup]';
	EXEC sp_executesql N'GRANT SELECT,INSERT,UPDATE,DELETE ON [ASRSysWorkflows] TO [ASRSysAdmins]';


	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbsys_mobileformlayout]') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[tbsys_mobileformlayout](
			[ID] [int] NOT NULL,
			[HeaderBackColor] [int] NOT NULL,
			[HeaderPictureID] [int] NULL,
			[HeaderPictureLocation] [tinyint] NOT NULL,
			[HeaderLogoID] [int] NULL,
			[HeaderLogoWidth] [int] NOT NULL,
			[HeaderLogoHeight] [int] NOT NULL,
			[HeaderLogoHorizontalOffset] [int] NOT NULL,
			[HeaderLogoVerticalOffset] [int] NOT NULL,
			[HeaderLogoHorizontalOffsetBehaviour] [tinyint] NOT NULL,
			[HeaderLogoVerticalOffsetBehaviour] [tinyint] NOT NULL,
			[MainBackColor] [int] NOT NULL,
			[MainPictureID] [int] NULL,
			[MainPictureLocation] [tinyint] NOT NULL,
			[FooterBackColor] [int] NOT NULL,
			[FooterPictureID] [int] NULL,
			[FooterPictureLocation] [tinyint] NOT NULL,
			[TodoTitleFontName] [varchar](255) NOT NULL,
			[TodoTitleFontSize] [float] NOT NULL,
			[TodoTitleFontBold] [bit] NOT NULL,
			[TodoTitleFontItalic] [bit] NOT NULL,
			[TodoDescFontName] [varchar](255) NOT NULL,
			[TodoDescFontSize] [float] NOT NULL,
			[TodoDescFontBold] [bit] NOT NULL,
			[TodoDescFontItalic] [bit] NOT NULL,
			[HomeItemFontName] [varchar](255) NOT NULL,
			[HomeItemFontSize] [float] NOT NULL,
			[HomeItemFontBold] [bit] NOT NULL,
			[HomeItemFontItalic] [bit] NOT NULL,
 		    CONSTRAINT [PK_tbsys_mobileformlayout] PRIMARY KEY CLUSTERED ([ID] ASC));';
 		 			   
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformlayout] ([ID], [HeaderBackColor], [HeaderPictureID], [HeaderPictureLocation], [HeaderLogoID], [HeaderLogoWidth], [HeaderLogoHeight], [HeaderLogoHorizontalOffset], [HeaderLogoVerticalOffset], [HeaderLogoHorizontalOffsetBehaviour], [HeaderLogoVerticalOffsetBehaviour], [MainBackColor], [MainPictureID], [MainPictureLocation], [FooterBackColor], [FooterPictureID], [FooterPictureLocation], [TodoTitleFontName], [TodoTitleFontSize], [TodoTitleFontBold], [TodoTitleFontItalic], [TodoDescFontName], [TodoDescFontSize], [TodoDescFontBold], [TodoDescFontItalic], [HomeItemFontName], [HomeItemFontSize], [HomeItemFontBold], [HomeItemFontItalic]) VALUES (1, 11829830, NULL, 5, NULL, 0, 0, 0, 0, 0, 0, 16777215, NULL, 5, 11829830, NULL, 5, N''Verdana'', 9.75, 1, 0, N''Verdana'', 8.25, 0, 0, N''Verdana'', 9.75, 1, 0);';
	END

	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbsys_mobileformelements]') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[tbsys_mobileformelements](
			[ID] [int] NOT NULL,
			[Form] [tinyint] NOT NULL,
			[Type] [tinyint] NOT NULL,
			[Name] [varchar](50) NULL,
			[Caption] [varchar](500) NULL,
			[FontName] [varchar](255) NULL,
			[FontSize] [float] NULL,
			[FontBold] [bit] NULL,
			[FontItalic] [bit] NULL,
			[Width] [int] NULL,
			[Height] [int] NULL,
			[BackStyle] [int] NULL,
			[BackColor] [int] NULL,
			[ForeColor] [int] NULL,
			[HorizontalOffset] [int] NULL,
			[VerticalOffset] [int] NULL,
			[HorizontalOffsetBehaviour] [tinyint] NULL,
			[VerticalOffsetBehaviour] [tinyint] NULL,
			[PasswordType] [bit] NULL,
			[PictureID] [int] NULL
		    CONSTRAINT [PK_tbsys_mobileformelements] PRIMARY KEY CLUSTERED ([ID] ASC))';

		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (7, 1, 2, N''lblUserName'', N''Username:'', N''Verdana'', 9.75, 1, 0, 90, 16, NULL, NULL, 0, 32, 110, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (8, 1, 3, N''txtUserName'', NULL, N''Verdana'', 9.75, 1, 0, 155, 99, NULL, 16777215, 0, 35, 110, 1, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (9, 1, 2, N''lblPassword'', N''Password:'', N''Verdana'', 9.75, 1, 0, 90, 16, NULL, NULL, 0, 32, 170, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (10, 1, 3, N''txtPassword'', NULL, N''Verdana'', 9.75, 1, 0, 155, 21, NULL, 16777215, 0, 35, 170, 1, 0, 1, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (11, 1, 2, N''lblRememberPwd'', N''Keep me signed in:'', N''Verdana'', 8.25, 0, 0, 145, 16, NULL, NULL, 0, 32, 224, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (12, 1, 0, N''btnLogin'', N''Sign In'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (13, 1, 0, N''btnForgotPwd'', N''Forgot Login Details'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (14, 1, 0, N''btnRegister'', N''New Registration'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (15, 2, 0, N''btnToDoList'', N''To Do List'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (17, 2, 0, N''btnLogout'', N''Sign Out'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (18, 2, 0, N''btnChangePwd'', N''Change Password'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (19, 3, 2, N''lblEmail'', N''Email address:'', N''Verdana'', 9.75, 1, 0, 110, 16, NULL, NULL, 0, 32, 110, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (20, 3, 3, N''txtEmail'', NULL, N''Verdana'', 9.75, 1, 0, 155, 21, NULL, 16777215, 0, 35, 110, 1, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (27, 3, 0, N''btnHome'', N''Home'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (28, 3, 0, N''btnRegister'', N''Register'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (29, 4, 2, N''lblCurrPassword'', N''Current Password:'', N''Verdana'', 9.75, 1, 0, 110, 16, NULL, NULL, 0, 32, 110, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (30, 4, 3, N''txtCurrPassword'', NULL, N''Verdana'', 9.75, 1, 0, 155, 21, NULL, 16777215, 0, 35, 110, 1, 0, 1, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (31, 4, 2, N''lblNewPassword'', N''New Password:'', N''Verdana'', 9.75, 1, 0, 110, 16, NULL, NULL, 0, 32, 170, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (32, 4, 3, N''txtNewPassword'', NULL, N''Verdana'', 9.75, 1, 0, 155, 21, NULL, 16777215, 0, 35, 170, 1, 0, 1, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (33, 4, 2, N''lblConfPassword'', N''Confirm Password:'', N''Verdana'', 9.75, 1, 0, 140, 16, NULL, NULL, 0, 35, 227, 1, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (34, 4, 3, N''txtConfPassword'', NULL, N''Verdana'', 9.75, 1, 0, 155, 21, NULL, 16777215, 0, 35, 227, 1, 0, 1, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (35, 4, 0, N''btnCancel'', N''Cancel'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (36, 4, 0, N''btnSubmit'', N''OK'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (37, 5, 0, N''btnCancel'', N''Cancel'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (38, 5, 0, N''btnRefresh'', N''Refresh'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (39, 6, 2, N''lblWelcome'', N''Enter your email address and an email will be sent to you confirming your login details.'', N''Verdana'', 8.25, 0, 0, 340, 32, NULL, NULL, 0, 32, 70, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (40, 6, 2, N''lblEmail'', N''Email address:'', N''Verdana'', 9.75, 1, 0, 110, 16, NULL, NULL, 0, 32, 110, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (41, 6, 3, N''txtEmail'', NULL, N''Verdana'', 9.75, 1, 0, 155, 21, NULL, 16777215, 0, 35, 110, 1, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (42, 6, 0, N''btnCancel'', N''Cancel'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (43, 6, 0, N''btnSubmit'', N''OK'', N''Verdana'', 6, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (44, 3, 2, N''lblWelcome'', N''Enter your registration details and an activation email will be sent to you.'', N''Verdana'', 8.25, 0, 0, 155, 32, NULL, NULL, 0, 50, 50, 1, 1, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (45, 4, 2, N''lblWelcome'', N''Enter your current and new passwords.'', N''Verdana'', 8.25, 0, 0, 340, 32, NULL, NULL, 0, 32, 70, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (46, 5, 2, N''lblNothingTodo'', N''You have nothing in your ''''to do'''' list.'', N''Verdana'', 8.25, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (47, 5, 2, N''lblInstruction'', N''Click on a ''''to do'''' item to view the details and complete your action.'', N''Verdana'', 8.25, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL);';	
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (48, 1, 2, N''lblWelcome'', N''Enter your login details and sign in.'', N''Verdana'', 8.25, 0, 0, 340, 32, NULL, NULL, 0, 32, 70, 0, 0, NULL, NULL);';
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [Width], [Height], [BackStyle], [BackColor], [ForeColor], [HorizontalOffset], [VerticalOffset], [HorizontalOffsetBehaviour], [VerticalOffsetBehaviour], [PasswordType], [PictureID]) VALUES (49, 2, 2, N''lblWelcome'', N''Click on an ''''item'''' to start an action.'', N''Verdana'', 8.25, 0, 0, 340, 32, NULL, NULL, 0, 32, 70, 0, 0, NULL, NULL);';

		DECLARE @maxid int, @signin int, @forgotlogin int, @newreg int,	@todolist int, @changepassword int, @signout int, @register int, @home int,	@ok int, @cancel int, @refresh int
		
		SELECT @maxid = MAX(PictureID) FROM dbo.ASRSysPictures
			
		SET @signin = null
		SET @forgotlogin = null
		SET @newreg = null
		SET @todolist = null
		SET @changepassword = null
		SET @signout = null
		SET @register = null
		SET @home = null
		SET @ok = null
		SET @cancel = null
		SET @refresh = null
	
		--TODO PG
		--INSERT INTO [dbo].[ASRSysPictures] VALUES (-1, '', null, 1)

		UPDATE tbsys_mobileformelements SET PictureID = @signin WHERE ID = 12
		UPDATE tbsys_mobileformelements SET PictureID = @forgotlogin WHERE ID = 13
		UPDATE tbsys_mobileformelements SET PictureID = @newreg WHERE ID = 14
		UPDATE tbsys_mobileformelements SET PictureID = @todolist WHERE ID = 15
		UPDATE tbsys_mobileformelements SET PictureID = @signout WHERE ID = 17
		UPDATE tbsys_mobileformelements SET PictureID = @changepassword WHERE ID = 18
		UPDATE tbsys_mobileformelements SET PictureID = @home WHERE ID = 27
		UPDATE tbsys_mobileformelements SET PictureID = @register WHERE ID = 28
		UPDATE tbsys_mobileformelements SET PictureID = @cancel WHERE ID = 35
		UPDATE tbsys_mobileformelements SET PictureID = @ok WHERE ID = 36
		UPDATE tbsys_mobileformelements SET PictureID = @cancel WHERE ID = 37
		UPDATE tbsys_mobileformelements SET PictureID = @refresh WHERE ID = 38
		UPDATE tbsys_mobileformelements SET PictureID = @cancel WHERE ID = 42
		UPDATE tbsys_mobileformelements SET PictureID = @ok WHERE ID = 43
	END
	UPDATE [dbo].[tbsys_mobileformelements] SET Caption = 'Enter your registration details and an activation email will be sent to you.' WHERE ID = 44;

	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbsys_mobilegroupworkflows]') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[tbsys_mobilegroupworkflows](
			[UserGroupID] [int] NOT NULL,
			[WorkflowID] [int] NOT NULL,
			[Pos] [int] NOT NULL,
		 CONSTRAINT [PK_tbsys_mobilegroupworkflows] PRIMARY KEY CLUSTERED ([UserGroupID] ASC, [WorkflowID] ASC));';
	END
		
	----------------------------------------------------------------------
	-- spASRMobileInstantiateWorkflow Stored Procedure
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRMobileInstantiateWorkflow]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRMobileInstantiateWorkflow];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRMobileInstantiateWorkflow]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;
	
	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRMobileInstantiateWorkflow]
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
		
		
	
	
	
	
	
	
	
/* ------------------------------------------------------------- */
PRINT 'Step 9 - System procedures'

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRDeleteRecord]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRDeleteRecord];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRInsertNewRecord]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRInsertNewRecord];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRRecordAmended]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRRecordAmended];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRUpdateRecord]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRUpdateRecord];
			
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRDeleteRecord]
(
    @piResult integer OUTPUT,   /* Output variable to hold the result. */
    @piTableID integer,			/* TableID being deleted from. */
    @psRealSource sysname,		/* RealSource being deleted from. */
    @piID integer				/* ID the record being deleted. */
)
AS
BEGIN
    SET NOCOUNT ON;

    /*  Delete the given record. */
    /* Check if the given record has been deleted or changed first. */
    /* Return 0 if the record was OK to delete. */
    /* NOT USED HERE - Return 1 if the record has been amended AND is still in the given table/view. */
    /* Return 2 if the record has been amended AND is no longer in the given table/view. */
    /* Return 3 if the record has been deleted from the table. */
    DECLARE @sSQL nvarchar(MAX),
		@psTableName sysname,
        @iCount integer;
    SET @piResult = 0;

	SELECT @psTableName = TableName FROM dbo.tbsys_tables WHERE TableID = @piTableID;

    /* Check that the record has not been updated by another user since it was last checked. */
    SET @sSQL = ''SELECT @piResult = COUNT(id)'' +
            '' FROM '' + @psTableName +
            '' WHERE id = '' + convert(varchar(MAX), @piID);
    EXECUTE sp_executesql @sSQL, N''@piResult int OUTPUT'', @iCount OUTPUT;  
    IF @iCount = 0
    BEGIN
        /* Record deleted. */
        SET @piResult = 3;
    END
    ELSE
    BEGIN
        /* Check if the record is still in the given realsource. */
        SET @sSQL = ''SELECT @piResult = COUNT(id)'' +
            '' FROM '' + @psRealSource +
            '' WHERE id = '' + convert(varchar(MAX), @piID);
        EXECUTE sp_executesql @sSQL, N''@piResult int OUTPUT'', @iCount OUTPUT;
        IF @iCount > 0
        BEGIN
            SET @sSQL = ''DELETE '' +
                '' FROM '' + @psRealSource +
                '' WHERE id = '' + convert(varchar(MAX), @piID);
            EXECUTE sp_executesql @sSQL;
        END
        ELSE
        BEGIN
            SET @piResult = 2;
        END
    END

END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRInsertNewRecord]
	(
		@piNewRecordID integer OUTPUT,   /* Output variable to hold the new record ID. */
		@psInsertString nvarchar(MAX)    /* SQL Insert string to insert the new record. */
	)
	AS
	BEGIN
		SET NOCOUNT ON;

		DECLARE @ssql nvarchar(MAX),
				@tablename varchar(255);

		-- Run the given SQL INSERT
		EXECUTE sp_executesql @psInsertString;

		-- Calculate the ID
		SET  @psInsertString = REPLACE('' '' + @psInsertString,'' INSERT INTO '','''')
		SET  @psInsertString = REPLACE('' '' + @psInsertString,'' INSERT '','''')
		SET @tablename = SUBSTRING(@psInsertString,0, CHARINDEX(''('', @psInsertString));

		SET @ssql = ''SELECT @ID = MAX(ID) FROM '' + @tablename;
		EXECUTE sp_executesql @ssql, N''@ID int OUTPUT'', @ID = @piNewRecordID OUTPUT;

END'


EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRRecordAmended]
(
    @piResult integer OUTPUT,	/* Output variable to hold the result. */
    @piTableID integer,			/* TableID being updated. */
    @psRealSource sysname,		/* RealSource being updated. */
    @piID integer,				/* ID the record being updated. */
    @piTimestamp integer		/* Original timestamp of the record being updated. */
)
AS
BEGIN
    /* Check if the given record has been deleted or changed by another user. */
    /* Return 0 if the record has NOT been amended. */
    /* Return 1 if the record has been amended AND is still in the given table/view. */
    /* Return 2 if the record has been amended AND is no longer in the given table/view. */
    /* Return 3 if the record has been deleted from the table. */
    SET NOCOUNT ON;
    DECLARE @iCurrentTimestamp integer,
        @sSQL nvarchar(MAX),
        @psTableName sysname,
        @iCount integer;
    SET @piResult = 0;

	SELECT @psTableName = TableName FROM dbo.tbsys_tables WHERE TableID = @piTableID;

    /* Check that the record has not been updated by another user since it was last checked. */
    SET @sSQL = ''SELECT @iCurrentTimestamp = convert(integer, timestamp)'' +
            '' FROM '' + @psTableName +
            '' WHERE id = '' + convert(varchar(MAX), @piID);
    EXECUTE sp_executesql @sSQL, N''@iCurrentTimestamp int OUTPUT'', @iCurrentTimestamp OUTPUT;
    
    IF @iCurrentTimestamp IS null
    BEGIN
        /* Record deleted. */
        SET @piResult = 3;
    END
    ELSE
    BEGIN
        IF @iCurrentTimestamp <> @piTimestamp
        BEGIN
            /* Record changed. Check if it is in the given realsource. */
           SET @sSQL = ''SELECT @piResult = COUNT(id)'' +
             '' FROM '' + @psRealSource +
             '' WHERE id = '' + convert(varchar(255), @piID);
           EXECUTE sp_executesql @sSQL, N''@piResult int OUTPUT'', @iCount OUTPUT;
           IF @iCount > 0
           BEGIN
               SET @piResult = 1;
           END
           ELSE
           BEGIN
               SET @piResult = 2;
           END
        END
    END
END'

EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRUpdateRecord]
(
    @piResult integer OUTPUT,		/* Output variable to hold the result. */
    @psUpdateString nvarchar(MAX),  /* SQL Update string to update the record. */
    @piTableID integer,				/* TableID being updated. */
    @psRealSource sysname,			/* RealSource being updated. */
    @piID integer,					/* ID the record being updated. */
    @piTimestamp integer			/* Original timestamp of the record being updated. */
)
AS
BEGIN
    SET NOCOUNT ON;

    /* Run the given SQL UPDATE string. */
    /* Check if the given record has been deleted or changed first. */
    /* Return 0 if the record was OK to update. */
    /* Return 1 if the record has been amended AND is still in the given table/view. */
    /* Return 2 if the record has been amended AND is no longer in the given table/view. */
    /* Return 3 if the record has been deleted from the table. */
    DECLARE @iCurrentTimestamp integer,
        @sSQL nvarchar(MAX),
        @psTableName sysname,
        @iCount integer;
    SET @piResult = 0;

	SELECT @psTableName = TableName FROM dbo.tbsys_tables WHERE TableID = @piTableID; 

    /* Check that the record has not been updated by another user since it was last checked. */
    SET @sSQL = ''SELECT @iCurrentTimestamp = convert(integer, timestamp)'' +
            '' FROM '' + @psTableName +
            '' WHERE id = '' + convert(varchar(MAX), @piID);
    EXECUTE sp_executesql @sSQL, N''@iCurrentTimestamp int OUTPUT'', @iCurrentTimestamp OUTPUT;
    
    IF @iCurrentTimestamp IS null
    BEGIN
        /* Record deleted. */
        SET @piResult = 3;
    END
    ELSE
    BEGIN
        IF (@iCurrentTimestamp <> @piTimestamp) AND (NOT @piTimestamp IS null)
        BEGIN
            /* Record changed. Check if it is in the given realsource. */
           SET @sSQL = ''SELECT @piResult = COUNT(id)'' +
             '' FROM '' + @psRealSource +
             '' WHERE id = '' + convert(varchar(255), @piID);
           EXECUTE sp_executesql @sSQL, N''@piResult int OUTPUT'', @iCount OUTPUT;
           IF @iCount > 0
           BEGIN
               SET @piResult = 1;
           END
           ELSE
           BEGIN
               SET @piResult = 2;
           END
        END
        ELSE
        BEGIN
            -- Run the given SQL UPDATE string.
            EXECUTE sp_executeSQL @psUpdateString;
        END
    END

END'




/* --------------------------------------------------- */
/* Remove unused stored procedures from the database.  */
/* --------------------------------------------------- */
DECLARE @dropsql nvarchar(max), @name nvarchar(max)

DECLARE c CURSOR FOR
SELECT ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES
WHERE ROUTINE_TYPE = 'PROCEDURE' AND 
(
ROUTINE_NAME LIKE 'sp_ASRDeleteRecord_%' OR
ROUTINE_NAME LIKE 'sp_ASRInsertNewRecord_%' OR
ROUTINE_NAME LIKE 'sp_ASRUpdateRecord_%' OR
ROUTINE_NAME LIKE 'spASRIntInsertNewRecord_%' OR
ROUTINE_NAME LIKE 'spASRIntUpdateRecord_%' OR
ROUTINE_NAME LIKE 'sp_ASRRecordAmended_%' OR
ROUTINE_NAME LIKE 'sp_ASRValidate[_][1-9]%'
);

OPEN c;
FETCH NEXT FROM c INTO @name;
WHILE @@FETCH_STATUS = 0
BEGIN
	SET @dropsql = 'DROP PROCEDURE [dbo].[' + @name + ']';

	EXECUTE sp_executesql @dropsql;
	
	FETCH NEXT FROM c INTO @name;
END
CLOSE c;
DEALLOCATE c;


/* ------------------------------------------------------------- */
PRINT 'Step 10 - Message Bus Integration'

	IF NOT EXISTS(SELECT * FROM sys.schemas where name = 'fusion')
		EXECUTE sp_executesql N'CREATE SCHEMA [fusion];';

	IF NOT EXISTS(SELECT * FROM sys.sysobjects where name = 'IdTranslation' AND xtype = 'U')
		EXECUTE sp_executesql N'CREATE TABLE [fusion].[IdTranslation](
			[TranslationName] [varchar](50) NOT NULL,
			[LocalId] [varchar](25) NOT NULL,
			[BusRef] [uniqueidentifier] NOT NULL);';

	IF NOT EXISTS(SELECT * FROM sys.sysobjects where name = 'MessageLog' AND xtype = 'U')
		EXECUTE sp_executesql N'CREATE TABLE [fusion].[MessageLog](
			[MessageType] [varchar](50) NOT NULL,
			[MessageRef] [uniqueidentifier] NOT NULL,
			[ReceivedDate] [datetime] NOT NULL,
			[Originator] [varchar](50) NULL);';

	IF NOT EXISTS(SELECT * FROM sys.sysobjects where name = 'MessageTracking' AND xtype = 'U')
		EXECUTE sp_executesql N'CREATE TABLE [fusion].[MessageTracking](
			[MessageType] [varchar](50) NOT NULL,
			[BusRef] [uniqueidentifier] NOT NULL,
			[LastGeneratedDate] [datetime] NULL,
			[LastProcessedDate] [datetime] NULL,
			[LastGeneratedXml] [varchar](max) NULL);';




	DECLARE	@perstableid integer,
			@columnid integer;

	-- Message table mapping defintion
	IF NOT EXISTS(SELECT * FROM sys.sysobjects where name = 'ASRSysFusionTypes' AND xtype = 'U')
	BEGIN
		EXECUTE sp_executesql N'CREATE TABLE [dbo].[ASRSysFusionTypes](
			[FusionTypeID] [int] NOT NULL,
			[FusionType] [nvarchar](40) NULL,
			[FilterID] [int] NULL,
			[ASRBaseTableID] [int] NULL,
			[IsVisible] [bit] NULL,
			[Version] numeric(5,2)
		    CONSTRAINT [PK_FusionTypeID] PRIMARY KEY NONCLUSTERED ([FusionTypeID] ASC));';

		SELECT @perstableid = [parametervalue] FROM dbo.[ASRSysModuleSetup] WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_TablePersonnel'
			AND [ParameterType] = 'PType_TableID';

		INSERT [dbo].[ASRSysFusionTypes] ([FusionTypeID], [FusionType], [ASRBaseTableID], [FilterID], [IsVisible], [Version]) VALUES (1, 'Staff Change', @perstableid, 0, 1, 1);

	END

	-- Message data mapping defintion
	IF NOT EXISTS(SELECT * FROM sys.sysobjects where name = 'ASRSysFusionFieldMappings' AND xtype = 'U')
	BEGIN
		EXECUTE sp_executesql N'CREATE TABLE [dbo].[ASRSysFusionFieldMappings](
			[FusionID] [int] NOT NULL,
			[NodeKey] [varchar](255) NOT NULL,
			[HRProValue] [varchar](100) NOT NULL,
			[FusionValue] [varchar](100) NOT NULL);';
	END


	-- Message column mapping defintion
	IF NOT EXISTS(SELECT * FROM sys.sysobjects where name = 'ASRSysFusionFieldDefinitions' AND xtype = 'U')
	BEGIN
		EXECUTE sp_executesql N'CREATE TABLE [dbo].[ASRSysFusionFieldDefinitions](
			[NodeKey] [varchar](255) NOT NULL,
			[FusionTypeID] [int] NOT NULL,
			[DataType] [tinyint] NOT NULL,
			[Mandatory] [bit] NOT NULL,
			[Description] [char](40) NOT NULL,
			[AlwaysTransfer] [bit] NULL,
			[IsKeyField] [bit] NULL,
			[IsCompanyCode] [bit] NULL,
			[IsEmployeeCode] [bit] NULL,
			[Direction] [int] NULL,
			[ASRMapType] [int] NULL,
			[ASRTableID] [int] NULL,
			[ASRColumnID] [int] NULL,
			[ASRExprID] [int] NULL,
			[ASRValue] [char](40) NULL,
			[ConvertData] [bit] NULL,
			[IsEmployeeName] [bit] NULL,
			[IsDepartmentCode] [bit] NULL,
			[IsDepartmentName] [bit] NULL,
			[IsFusionCode] [bit] NULL,
			[PreventModify] [bit] NULL);';

		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'TITLE', 0, 0, 0, 1, 'Title',0,'')

		SELECT @columnid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup] WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsForename' AND [ParameterType] = 'PType_ColumnID';
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'FORENAMES', 0, @perstableid, @columnid, 1, 'Forenames',0,'')

		SELECT @columnid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup] WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsSurname' AND [ParameterType] = 'PType_ColumnID';
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'SURNAME', 0, @perstableid, @columnid, 1, 'Surname',0,'')

		SELECT @columnid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup] WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsSSIWelcome' AND [ParameterType] = 'PType_ColumnID';
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'PREFERREDNAME', 0, @perstableid, @columnid, 1, 'Preferred Name',0,'')

		SELECT @columnid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup] WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsDateOfBirth' AND [ParameterType] = 'PType_ColumnID';
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,11,'DOB', 0, @perstableid, @columnid, 1, 'Date of Birth',0,'')

		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'EMPLOYEETYPE', 0, 0, 0, 1, 'Employee Type',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'EMPLOYEESTATUS', 0, 0, 0, 1, 'Employee Status',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'HOMEPHONENUMBER', 0, 0, 0, 1, 'Home Phone Number',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'WORKMOBILE', 0, 0, 0, 1, 'Work Mobile',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'PERSONALMOBILE', 0, 0, 0, 1, 'Personal Mobile',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'EMAIL', 0, 0, 0, 1, 'Email',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'PERSONALEMAIL', 0, 0, 0, 1, 'Personal Email',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'ADDRESSLINE1', 0, 0, 0, 1, 'Address Line 1',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'ADDRESSLINE2', 0, 0, 0, 1, 'Address Line 2',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'ADDRESSLINE3', 0, 0, 0, 1, 'Address Line 3',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'ADDRESSLINE4', 0, 0, 0, 1, 'Address Line 4',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'ADDRESSLINE5', 0, 0, 0, 1, 'Address Line 5',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'POSTCODE', 0, 0, 0, 1, 'Postcode',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'GENDER', 0, 0, 0, 1, 'Gender',0,'')

		SELECT @columnid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup] WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsStartDate' AND [ParameterType] = 'PType_ColumnID';
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,11,'STARTDATE', 0, @perstableid, @columnid, 1, 'Start Date',0,'')

		SELECT @columnid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup] WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsLeavingDate' AND [ParameterType] = 'PType_ColumnID';
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,11,'LEAVINGDATE', 0, @perstableid, @columnid, 1, 'Leaving Date',0,'')

		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'LEAVINGREASON', 0, 0, 0, 1, 'Leaving Reason',0,'')
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'COMPANYNAME', 0, 0, 0, 1, 'Company Name',0,'')

		SELECT @columnid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup] WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsJobTitle' AND [ParameterType] = 'PType_ColumnID';
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'jobtitle', 0, @perstableid, @columnid, 1, 'Job Title',0,'')

		SELECT @columnid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup] WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_FieldsManagerStaffNo' AND [ParameterType] = 'PType_ColumnID';
		INSERT [dbo].[ASRSysFusionFieldDefinitions] ([FusionTypeID], [DataType], [NodeKey], [ASRMapType], [ASRTableID], [ASRColumnID], [Mandatory], [Description], [ASRExprID], [ASRValue]) VALUES (1,12,'MANAGERREF', 0, @perstableid, @columnid, 1, 'Manager',0,'')

	END


	-- Configure the service broker
	IF NOT EXISTS(SELECT name FROM sys.service_message_types WHERE name = 'TriggerFusionSend')
		EXECUTE sp_executesql N'CREATE MESSAGE TYPE TriggerFusionSend VALIDATION = NONE;';

	IF NOT EXISTS(SELECT name FROM sys.service_contracts WHERE name = 'TriggerFusionContract')
		EXECUTE sp_executesql N'CREATE CONTRACT TriggerFusionContract (TriggerFusionSend SENT BY INITIATOR);';

	IF NOT EXISTS(SELECT name FROM sys.service_queues WHERE name = 'qFusion')
		EXECUTE sp_executesql N'CREATE QUEUE fusion.qFusion WITH STATUS = ON;';

	IF NOT EXISTS(SELECT name FROM sys.services WHERE name = 'FusionApplicationService')
		EXECUTE sp_executesql N'CREATE SERVICE FusionApplicationService ON QUEUE fusion.qFusion (TriggerFusionContract);';

	IF NOT EXISTS(SELECT name FROM sys.services WHERE name = 'FusionConnectorService')
		EXECUTE sp_executesql N'CREATE SERVICE FusionConnectorService ON QUEUE fusion.qFusion (TriggerFusionContract);';


	-- Apply the stored procedures
	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spSendMessage]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spSendMessage];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spSendMessageCheckContext]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spSendMessageCheckContext];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spMessageTrackingSetLastProcessedDate]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spMessageTrackingSetLastProcessedDate];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spMessageTrackingSetLastGeneratedXml]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spMessageTrackingSetLastGeneratedXml];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spMessageTrackingSetLastGeneratedDate]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spMessageTrackingSetLastGeneratedDate];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spMessageTrackingGetLastMessageDates]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spMessageTrackingGetLastMessageDates];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spMessageTrackingGetLastGeneratedXml]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spMessageTrackingGetLastGeneratedXml];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spMessageLogCheck]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spMessageLogCheck];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spMessageLogAdd]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spMessageLogAdd];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spIdTranslateSetBusRef]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spIdTranslateSetBusRef];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spIdTranslateGetLocalId]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spIdTranslateGetLocalId];

	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spIdTranslateGetBusRef]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spIdTranslateGetBusRef];



	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spIdTranslateGetBusRef
	--
	-- Purpose: Converts a local identifier into a uniqueidentifier for the bus, 
	--			returning consistent value for all future conversions.  
	--          This will create a new identifier where one is not found where
	--			@CanGenerate = 1
	--
	-- Returns: 0 = success, 1 = failure
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spIdTranslateGetBusRef]
		(
			@TranslationName varchar(50),
			@LocalId varchar(25),
			@BusRef uniqueidentifier output,
			@CanGenerate bit = 1
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		SET @BusRef = NULL;
	
		SELECT @BusRef = BusRef from [fusion].IdTranslation WITH (ROWLOCK) 
			WHERE TranslationName = @TranslationName AND LocalId = @LocalId;
	
		IF @@ROWCOUNT = 0
		BEGIN
			IF @CanGenerate = 1
			BEGIN
				SET @BusRef = NEWID();
			
				INSERT [fusion].IdTranslation WITH (ROWLOCK) (TranslationName, LocalId, BusRef) 
						VALUES (@TranslationName, @LocalId, @BusRef);
					
				RETURN 0;
			END
			RETURN 1;
		END

		RETURN 0;
	END';


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spIdTranslateGetLocalId
	--
	-- Purpose: Finds the local id equivelant for the given Bus reference number, 
	--          assuming it has previous been created through spIdTranslateSetBusRef
	--
	-- Returns: 
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spIdTranslateGetLocalId]
		(
			@TranslationName varchar(50),
			@BusRef uniqueidentifier,
			@LocalId varchar(25) output
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		SET @LocalId = null;
	
		SELECT @LocalId = LocalId from [fusion].IdTranslation WITH (ROWLOCK) 
			WHERE TranslationName = @TranslationName and BusRef = @BusRef;
	END';


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spIdTranslateSetBusRef
	--
	-- Purpose: Sets the conversion of a given local reference into the given bus ref
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spIdTranslateSetBusRef]
		(
			@TranslationName varchar(50),
			@LocalId varchar(25),
			@BusRef uniqueidentifier
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		BEGIN TRAN;
	
		DELETE fusion.IdTranslation WITH (ROWLOCK) 
			WHERE TranslationName = @TranslationName and LocalId = @LocalId;
		
		INSERT fusion.IdTranslation WITH (ROWLOCK) (TranslationName, LocalId, BusRef) 
			VALUES (@TranslationName, @LocalId, @BusRef);

		COMMIT TRAN;
	END	'


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spMessageLogAdd
	--
	-- Purpose: Adds fact that message has been processed to local message log
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spMessageLogAdd]
		(
			@MessageType varchar(50),
			@MessageRef uniqueidentifier,
			@Originator varchar(50) = NULL
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		INSERT fusion.MessageLog (MessageType, MessageRef, Originator, ReceivedDate) VALUES (@MessageType, @MessageRef, @Originator, GETUTCDATE());

	END'


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spMessageLogCheck
	--
	-- Purpose: Checks whether message has been processed before
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spMessageLogCheck]
		(
			@MessageType varchar(50),
			@MessageRef uniqueidentifier,
			@ReceivedBefore bit output
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		IF EXISTS ( SELECT * FROM fusion.MessageLog WHERE MessageType = @MessageType AND MessageRef = @MessageRef )
		BEGIN
			SET @ReceivedBefore = 1
		END
		ELSE
		BEGIN
			SET @ReceivedBefore = 0
		END
	END'


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spMessageTrackingGetLastGeneratedXml
	--
	-- Purpose: Gets the last generated XML for a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spMessageTrackingGetLastGeneratedXml]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		SELECT LastGeneratedXml
			FROM fusion.MessageTracking
			WHERE MessageType = @MessageType AND BusRef = @BusRef;

	END'



	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spMessageTrackingGetLastMessageDates
	--
	-- Purpose: Gets the last processing date of a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spMessageTrackingGetLastMessageDates]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier
		)

	AS
	BEGIN
		SET NOCOUNT ON;
	
		SELECT LastProcessedDate, LastGeneratedDate
			FROM fusion.MessageTracking
			WHERE MessageType = @MessageType AND BusRef = @BusRef;

	END'


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spMessageTrackingSetLastGeneratedDate
	--
	-- Purpose: Sets the last processed date of a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spMessageTrackingSetLastGeneratedDate]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier,
			@LastGeneratedDate datetime
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		IF EXISTS (SELECT * FROM [fusion].MessageTracking
				   WHERE MessageType = @MessageType AND BusRef = @BusRef)
		BEGIN	
			UPDATE [fusion].MessageTracking
			   SET LastGeneratedDate = @LastGeneratedDate
			   WHERE MessageType = @MessageType AND BusRef = @BusRef
		END
		ELSE
		BEGIN
			INSERT [fusion].MessageTracking (MessageType, BusRef, LastGeneratedDate)
				VALUES (@MessageType, @BusRef, @LastGeneratedDate)
		END		
	END'


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spMessageTrackingSetLastGeneratedXml
	--
	-- Purpose: Sets the last generated XML for a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spMessageTrackingSetLastGeneratedXml]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier,
			@LastGeneratedXml varchar(max)
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		IF EXISTS (SELECT * FROM fusion.MessageTracking
				   WHERE MessageType = @MessageType AND BusRef = @BusRef)
		BEGIN	
			UPDATE fusion.MessageTracking
			   SET LastGeneratedXml = @LastGeneratedXml
			   WHERE MessageType = @MessageType AND BusRef = @BusRef
		END
		ELSE
		BEGIN
			INSERT fusion.MessageTracking (MessageType, BusRef, LastGeneratedXml)
				VALUES (@MessageType, @BusRef, @LastGeneratedXml)
		END		
	END'


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spMessageTrackingSetLastProcessedDate
	--
	-- Purpose: Sets the last processed date of a given message
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spMessageTrackingSetLastProcessedDate]
		(
			@MessageType varchar(50),
			@BusRef uniqueidentifier,
			@LastProcessedDate datetime
		)

	AS
	BEGIN
		SET NOCOUNT ON;
		
		IF EXISTS (SELECT * FROM fusion.MessageTracking
				   WHERE MessageType = @MessageType AND BusRef = @BusRef)
		BEGIN	
			UPDATE fusion.MessageTracking
			   SET LastProcessedDate = @LastProcessedDate
			   WHERE MessageType = @MessageType AND BusRef = @BusRef
		END
		ELSE
		BEGIN
			INSERT fusion.MessageTracking (MessageType, BusRef, LastProcessedDate)
				VALUES (@MessageType, @BusRef, @LastProcessedDate)
		END		
	END'


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spSendMessage
	--
	-- Purpose: Triggers a message to be sent
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spSendMessage]
		(
			@MessageType varchar(50),
			@LocalId int
		)
	AS
	BEGIN
		SET NOCOUNT ON;
	
		DECLARE @DialogHandle uniqueidentifier;
		SET @DialogHandle = NEWID();

		BEGIN DIALOG @DialogHandle 
			FROM SERVICE MessageApplicationService 
			TO SERVICE ''MessageConnectorService''
			ON CONTRACT TriggerMessageContract
			WITH ENCRYPTION = OFF;
		
		DECLARE @msg varchar(max);

		SET @msg = (SELECT	@MessageType AS MessageType, 
							@LocalId as LocalId,
							CONVERT(varchar(50), GETUTCDATE(), 126)+''Z'' as TriggerDate 
						FOR XML PATH(''SendMessage''));	
		
		SEND ON CONVERSATION @DialogHandle
			MESSAGE TYPE TriggerMessageSend (@msg);
	 
		END CONVERSATION @DialogHandle;

	END'


	EXECUTE sp_executesql N'
	---------------------------------------------------------------------------------
	-- Name:    spSendMessageCheckContext
	--
	-- Purpose: Triggers a message to be sent, checking context
	--          to see if we are in the process of updating according to
	--          this same message being received (preventing multi-master
	--          re-publish scenario)
	--
	-- Returns: n/a
	---------------------------------------------------------------------------------

	CREATE PROCEDURE [fusion].[spSendMessageCheckContext]
		(
			@MessageType varchar(50),
			@LocalId int
		)
	AS
	BEGIN
		SET NOCOUNT ON;
	
		DECLARE @ContextInfo varbinary(128);
 
		SELECT @ContextInfo = CAST( ''Fusion:''+@MessageType AS VARBINARY(128) );
 
		IF CONTEXT_INFO() IS NULL OR CONTEXT_INFO() <> @ContextInfo
		BEGIN	
			EXEC fusion.spSendMessage @MessageType, @LocalId;
		END
	END'


/* ------------------------------------------------------------- */
/* Update various system calculation. */
/* ------------------------------------------------------------- */
PRINT 'Step 11 - System Calculations'

	DELETE FROM dbo.[tbstat_componentcode] WHERE [ID] = 4 AND [isoperator] = 1;
	INSERT [dbo].[tbstat_componentcode] ([id], [objectid], [code], [datatype], [name], [isoperator], [operatortype], [aftercode]) 
		VALUES (4, 'a34f7387-91a1-40d6-b42f-f8032609cfd6', '/ NULLIF(', NULL, 'Divided by', 1, 0, ',0)');

	UPDATE dbo.[tbstat_componentcode] SET [recordidrequired] = 1 WHERE [ID] = 43 AND [isoperator] = 0;



/* ------------------------------------------------------------- */
/* Update the database version flag in the ASRSysSettings table. */
/* Dont Set the flag to refresh the stored procedures            */
/* ------------------------------------------------------------- */
PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '5.0';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v5.0')


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

SELECT @NVarCommand = 'USE ['+@DBName + ']
GRANT VIEW DEFINITION TO public'
EXEC sp_executesql @NVarCommand


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v5.0 Of OpenHR'
