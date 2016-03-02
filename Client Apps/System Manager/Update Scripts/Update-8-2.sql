/* --------------------------------------------------- */
/* Update the database from version 8.1 to version 8.2 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(MAX),
	@iSQLVersion int,
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
IF (@sDBVersion <> '8.1') and (@sDBVersion <> '8.2')
BEGIN
	RAISERROR('The current database version is incompatible with this update script', 16, 1)
	RETURN
END

-- Only allow script to be run on SQL2008 or above
SELECT @iSQLVersion = convert(float,substring(@@version,charindex('-',@@version)+2,2))
IF (@iSQLVersion < 10)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of OpenHR', 16, 1)
	RETURN
END


/* ------------------------------------------------------- */
PRINT 'Step - Mail Merge additions'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'UploadTemplate')
		EXEC sp_executesql N'ALTER TABLE ASRSysMailMergeName ADD UploadTemplate varbinary(MAX) NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'UploadTemplateName')
		EXEC sp_executesql N'ALTER TABLE ASRSysMailMergeName ADD UploadTemplateName nvarchar(255) NULL;';


/* ------------------------------------------------------- */
PRINT 'Step - Audit Changes'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_fieldlastchangedate]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_fieldlastchangedate];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_fieldchangedbetweentwodates]')AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_fieldchangedbetweentwodates];

	IF EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysAuditTrail', 'U') AND name = 'TableID')
		EXEC sp_executesql N'ALTER TABLE ASRSysAuditTrail ALTER COLUMN tableid integer NULL'

	ALTER TABLE [ASRSysAuditAccess] ALTER COLUMN [ComputerName] varchar(255);


/* ------------------------------------------------------- */
PRINT 'Step - Calculation Changes'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT * FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsys_weekdaysbetweentwodates]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsys_weekdaysbetweentwodates];

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfsys_weekdaysbetweentwodates](
			@datefrom AS datetime,
			@dateto AS datetime)
		RETURNS integer
		WITH SCHEMABINDING
		AS
		BEGIN
	
			DECLARE @result integer;

			SELECT @result = CASE 
				WHEN DATEDIFF (day, @datefrom, @dateto) < 0 THEN 0
				ELSE (DATEDIFF(dd, @datefrom, @dateto) + 1)
					- (DATEDIFF(wk, @datefrom, @dateto) * 2)
					- (CASE WHEN DATEPART(dw, @datefrom) = 1 THEN 1 ELSE 0 END)
					- (CASE WHEN DATEPART(dw, @dateto) = 7 THEN 1 ELSE 0 END)
					END;
				
			RETURN ISNULL(@result,0);
		
		END'

	-- Previously dropped above
	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfsys_fieldlastchangedate](
			@colrefID	varchar(32),
			@recordID	integer
		)
		RETURNS datetime
		WITH SCHEMABINDING
		AS
		BEGIN

			DECLARE @result		datetime,
					@tableid	integer,
					@columnid	integer;
		
			SET @tableid = SUBSTRING(@colrefID, 1, 8);
			SET @columnid = SUBSTRING(@colrefID, 10, 8);

			SELECT TOP 1 @result = DATEADD(dd, 0, DATEDIFF(dd, 0, [DateTimeStamp])) FROM dbo.[ASRSysAuditTrail]
				WHERE [ColumnID] = @columnid AND [TableID] = @tableID
					AND @recordID = [RecordID]
				ORDER BY [DateTimeStamp] DESC ;

			RETURN @result;

		END'

	EXEC sp_executesql N'CREATE FUNCTION [dbo].[udfsys_fieldchangedbetweentwodates](
		@colrefID	varchar(32),
		@fromdate	datetime,
		@todate		datetime,
		@recordID	integer
	)
	RETURNS bit
	WITH SCHEMABINDING
	AS
	BEGIN

		DECLARE @result		bit,
				@tableid	integer,
				@columnid	integer;
		
		SET @tableid = SUBSTRING(@colrefID, 1, 8);
		SET @columnid = SUBSTRING(@colrefID, 10, 8);
		SET @fromdate = DATEADD(dd, 0, DATEDIFF(dd, 0, @fromdate));
		SET @todate = DATEADD(dd, 0, DATEDIFF(dd, 0, @todate));

		SELECT @result = CASE WHEN
				EXISTS(SELECT [DateTimeStamp] FROM dbo.[ASRSysAuditTrail]
					WHERE [ColumnID] = @columnid AND [TableID] = @tableID
					AND @recordID = [RecordID] 
					AND [DateTimeStamp] >= @fromdate AND DateTimeStamp < @todate + 1)
				THEN 1 ELSE 0 END;

		RETURN @result;

	END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[udfsysStringToTable]') AND xtype in (N'FN', N'IF', N'TF'))
		DROP FUNCTION [dbo].[udfsysStringToTable];

	EXEC sp_executesql N'CREATE FUNCTION dbo.[udfsysStringToTable] (           
		  @String nvarchar(MAX),
		  @delimiter nvarchar(2))
	RETURNS @Table TABLE( Splitcolumn nvarchar(MAX)) 
	BEGIN

		DECLARE @Xml AS XML;
		SET @Xml = cast((''<A>''+replace(@String,@delimiter,''</A><A>'')+''</A>'') AS XML);

		INSERT INTO @Table SELECT LTRIM(RTRIM(A.value(''.'', ''nvarchar(max)''))) AS [Column] FROM @Xml.nodes(''A'') AS FN(A);
		RETURN;

	END'


/* ------------------------------------------------------- */
PRINT 'Step - Export additions'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'TransformFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD TransformFile nvarchar(MAX) NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'XMLDataNodeName')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD XMLDataNodeName nvarchar(50) NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'LastSuccessfulOutput')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD LastSuccessfulOutput datetime NULL;';
		
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'AuditChangesOnly')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD AuditChangesOnly bit NULL;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'XSDFileName')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD XSDFileName nvarchar(255) NULL, PreserveTransformPath bit, PreserveXSDPath bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitXMLNodesFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitXMLNodesFile bit;';

	ALTER TABLE [ASRSysExportName] ALTER COLUMN [HeaderText] varchar(MAX);
	ALTER TABLE [ASRSysExportName] ALTER COLUMN [FooterText] varchar(MAX);

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'StripDelimiterFromData')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD StripDelimiterFromData bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitFile')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitFile bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'SplitFileSize')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD SplitFileSize int;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'LinkedServer')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD LinkedServer nvarchar(255);';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'LinkedCatalog')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD LinkedCatalog nvarchar(255);';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'LinkedTable')
		EXEC sp_executesql N'ALTER TABLE ASRSysExportName ADD LinkedTable nvarchar(255);';


/* ------------------------------------------------------- */
PRINT 'Step - Import additions'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysImportName', 'U') AND name = 'LinkedServer')
		EXEC sp_executesql N'ALTER TABLE ASRSysImportName ADD LinkedServer nvarchar(255);';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysImportName', 'U') AND name = 'LinkedCatalog')
		EXEC sp_executesql N'ALTER TABLE ASRSysImportName ADD LinkedCatalog nvarchar(255);';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysImportName', 'U') AND name = 'LinkedTable')
		EXEC sp_executesql N'ALTER TABLE ASRSysImportName ADD LinkedTable nvarchar(255);';


/* ------------------------------------------------------- */
PRINT 'Step - Workflow additions'
/* ------------------------------------------------------- */

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowElements', 'U') AND name = 'RequiresAuthentication')
		EXEC sp_executesql N'ALTER TABLE ASRSysWorkflowElements ADD RequiresAuthentication bit NULL;';

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
		
			SELECT @iSQLVersion = convert(float,substring(@@version,charindex(''-'',@@version)+2,2));
		
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
		
		END';

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
		
			SELECT @iSQLVersion = convert(float,substring(@@version,charindex(''-'',@@version)+2,2));
		
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
		
		END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRInstantiateTriggeredWorkflows]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows];
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows]
		AS
		BEGIN
			DECLARE
				@iQueueID			integer,
				@iWorkflowID		integer,
				@iRecordID			integer,
				@iInstanceID		integer,
				@iStartElementID	integer,
				@iTemp				integer,
				@iBaseTable		integer,
				@iParent1TableID	integer,
				@iParent1RecordID	integer,
				@iParent2TableID	integer,
				@iParent2RecordID	integer,
				@TargetName varchar(MAX);

			DECLARE @succeedingElements table(elementID int);
			DECLARE	@outputTable table (id int NOT NULL);
		
			DECLARE triggeredWFCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT Q.queueID,
				Q.recordID,
				TL.workflowID,
				Q.parent1TableID,
				Q.parent1RecordID,
				Q.parent2TableID,
				Q.parent2RecordID,
				WF.baseTable
			FROM ASRSysWorkflowQueue Q
			INNER JOIN ASRSysWorkflowTriggeredLinks TL ON Q.linkID = TL.linkID
			INNER JOIN ASRSysWorkflows WF ON TL.workflowID = WF.ID
				AND WF.enabled = 1
			WHERE Q.dateInitiated IS null
				AND datediff(dd,DateDue,getdate()) >= 0
		
			OPEN triggeredWFCursor
			FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID, @iBaseTable
			WHILE (@@fetch_status = 0) 
			BEGIN
				UPDATE ASRSysWorkflowQueue
				SET dateInitiated = getDate()
				WHERE queueID = @iQueueID;

				EXEC [dbo].[sp_ASRIntGetRecordDescription] @iBaseTable, @iRecordID, 0, 0, @TargetName OUTPUT;
				
				-- Create the Workflow Instance record, and remember the ID. */
				INSERT INTO ASRSysWorkflowInstances (workflowID, 
					initiatorID, 
					status, 
					userName, 
					parent1TableID,
					parent1RecordID,
					parent2TableID,
					parent2RecordID,
					pageno,
					TargetName)
				OUTPUT inserted.ID INTO @outputTable
				VALUES (@iWorkflowID, 
					@iRecordID, 
					0, 
					''<Triggered>'',
					@iParent1TableID,
					@iParent1RecordID,
					@iParent2TableID,
					@iParent2RecordID,
					0,
					@TargetName)
								
				SELECT @iInstanceID = id FROM @outputTable;
				
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
						OR ASRSysWorkflowElementItems.itemType = 7
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
				
				FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID, @iBaseTable
			END
			CLOSE triggeredWFCursor
			DEALLOCATE triggeredWFCursor
		END';

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements];
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



/* ------------------------------------------------------- */
PRINT 'Step - Branding'
/* ------------------------------------------------------- */

	EXEC sp_executesql N'UPDATE ASRSysPermissionCategories SET [description] = ''OpenHR Web'' WHERE categoryID = 19';
	EXEC sp_executesql N'UPDATE ASRSysPermissionItems SET [description] = ''OpenHR Web'' WHERE itemID = 4';
	EXEC spsys_setsystemsetting 'support', 'email', 'ohrsupport@advancedcomputersoftware.com';


/* ------------------------------------------------------- */
PRINT 'Step - Database Metadata locking'
/* ------------------------------------------------------- */

IF EXISTS(SELECT * FROM sys.objects WHERE name = 'ASRSysTables' AND type = 'V')
BEGIN
	EXECUTE sp_executeSQL N'DROP VIEW [ASRSysTables]';

	EXEC sp_rename 'tbsys_tables', 'ASRSysTables';

	EXECUTE sp_executeSQL N'ALTER TABLE ASRSysTables
		ADD [Guid] uniqueidentifier,
			[Locked] bit';

	EXECUTE sp_executeSQL N'MERGE INTO ASRSysTables t
		USING tbsys_scriptedobjects o ON o.ObjectType = 1 AND o.[targetid] = t.tableID
		WHEN MATCHED THEN
		UPDATE 
			SET [Locked] = o.[Locked], [Guid] = o.[Guid];';

END

IF EXISTS(SELECT * FROM sys.objects WHERE name = 'ASRSysColumns' AND type = 'V')
BEGIN
	EXECUTE sp_executeSQL N'DROP VIEW [ASRSysColumns]';

	EXEC sp_rename 'tbsys_columns', 'ASRSysColumns';

	EXECUTE sp_executeSQL N'ALTER TABLE ASRSysColumns
		ADD [Guid] uniqueidentifier,
			[Locked] bit';

	EXECUTE sp_executeSQL N'MERGE INTO ASRSysColumns c
		USING tbsys_scriptedobjects o ON o.ObjectType = 2 AND o.[targetid] = c.ColumnID
		WHEN MATCHED THEN
		UPDATE 
			SET [Locked] = o.[Locked], [Guid] = o.[Guid];';

END


IF EXISTS(SELECT * FROM sys.objects WHERE name = 'ASRSysViews' AND type = 'V')
BEGIN
	EXECUTE sp_executeSQL N'DROP VIEW [ASRSysViews]';

	EXEC sp_rename 'tbsys_views', 'ASRSysViews';

	EXECUTE sp_executeSQL N'ALTER TABLE ASRSysViews
		ADD [Guid] uniqueidentifier,
			[Locked] bit';

	EXECUTE sp_executeSQL N'MERGE INTO ASRSysViews v
		USING tbsys_scriptedobjects o ON o.ObjectType = 3 AND o.[targetid] = v.ViewId
		WHEN MATCHED THEN
		UPDATE 
			SET [Locked] = o.[Locked], [Guid] = o.[Guid];';

END


IF EXISTS(SELECT * FROM sys.objects WHERE name = 'ASRSysWorkflows' AND type = 'V')
BEGIN
	EXECUTE sp_executeSQL N'DROP VIEW [ASRSysWorkflows]';

	EXEC sp_rename 'tbsys_workflows', 'ASRSysWorkflows';

	EXECUTE sp_executeSQL N'ALTER TABLE ASRSysWorkflows
		ADD [Guid] uniqueidentifier,
			[Locked] bit';

	EXECUTE sp_executeSQL N'MERGE INTO ASRSysWorkflows w
		USING tbsys_scriptedobjects o ON o.ObjectType = 10 AND o.[targetid] = w.Id
		WHEN MATCHED THEN
		UPDATE 
			SET [Locked] = o.[Locked], [Guid] = o.[Guid];';

END

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRRecordAmended]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRRecordAmended];


	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[sp_ASRRecordAmended]
	(
		@piResult integer OUTPUT,	/* Output variable to hold the result. */
		@piTableID integer,			/* TableID being updated. */
		@psRealSource sysname,		/* RealSource being updated. */
		@piID integer,				/* ID the record being updated. */
		@piTimestamp integer		/* Original timestamp of the record being updated. */
	)
	WITH EXECUTE AS ''dbo''
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

		SELECT @psTableName = TableName FROM ASRSysTables WHERE TableID = @piTableID;

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

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spadmin_optimiserecordsave]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spadmin_optimiserecordsave];

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spadmin_optimiserecordsave]
		AS
		BEGIN

			SET NOCOUNT ON;

			DECLARE @sCode nvarchar(MAX);

			SET @sCode = '''';
			SELECT @sCode = @sCode + ''UPDATE dbo.[tbuser_'' + [tablename] + ''] SET [updflag] = 0 WHERE [id] = 0;'' + CHAR(13)
				FROM ASRSysTables
				WHERE [TableType] IN (1,2)
				ORDER BY [tabletype];

			EXECUTE sp_executesql @sCode;
	
		END';

/* ------------------------------------------------------- */
PRINT 'Step - Database Hardening'
/* ------------------------------------------------------- */

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRWorkflowValidateService]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRWorkflowValidateService];
	EXEC sp_executesql N'CREATE PROCEDURE [dbo].[spASRWorkflowValidateService](@allow bit OUTPUT)
		WITH ENCRYPTION
		AS
		BEGIN

			DECLARE @iUserGroupID		integer,
				@sUserGroupName			sysname,
				@fSysSecMgr				bit,
				@sActualUserName		sysname;

			EXEC [dbo].[spASRIntGetActualUserDetails]
				@sActualUserName OUTPUT,
				@sUserGroupName OUTPUT,
				@iUserGroupID OUTPUT;

			SELECT @allow = CASE WHEN count(*) > 0 THEN 1 ELSE 0 END
				FROM ASRSysGroupPermissions gp
					INNER JOIN ASRSysPermissionItems pi ON gp.itemID = pi.itemID
					INNER JOIN ASRSysPermissionCategories pc ON pi.categoryID = pc.categoryID
					INNER JOIN sys.database_principals u ON gp.groupName = u.name
				WHERE u.principal_id = @iUserGroupID
					AND pi.itemKey = ''SYSTEMMANAGER'' AND gp.permitted = 1 AND pc.categorykey = ''MODULEACCESS'';

		END'


	EXECUTE sp_executeSQL N'CREATE PROCEDURE #spASRTempHardenTables(@tablename nvarchar(MAX), @adminrole nvarchar(MAX), @allusersHaveSelect bit)
	AS
	BEGIN

		DECLARE @NVarCommand nvarchar(MAX) = '''';

		SELECT @NVarCommand = @NVarCommand + ''REVOKE SELECT, UPDATE, INSERT, DELETE ON ['' + @tablename + ''] TO ['' + U.name + ''];''
			FROM sys.database_permissions P 
			JOIN sys.tables T ON P.major_id = T.object_id 
			JOIN sysusers U ON U.uid = P.grantee_principal_id
			WHERE t.name = @tablename;			
		EXECUTE sp_executeSQL @NVarCommand;

		SET @NVarCommand = ''GRANT SELECT, INSERT, UPDATE, DELETE ON ['' + @tablename + ''] TO ['' + @adminrole + ''];'';
		EXECUTE sp_executeSQL @NVarCommand;	

		IF @allusersHaveSelect = 1
		BEGIN
			SET @NVarCommand = ''GRANT SELECT ON ['' + @tablename + ''] TO [ASRSysGroup];'';
			EXECUTE sp_executeSQL @NVarCommand;
		END
	END';


	IF EXISTS (SELECT * FROM sys.database_principals WHERE name = N'ASRSysAdmins' AND type = 'R')
	BEGIN
		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand +  'EXEC sp_droprolemember @rolename = [ASRSysAdmins], @membername = [' + members.[name] + '];'
			FROM sys.database_role_members AS rolemembers
				JOIN sys.database_principals AS roles ON roles.[principal_id] = rolemembers.[role_principal_id]
				JOIN sys.database_principals AS members ON members.[principal_id] = rolemembers.[member_principal_id]
			WHERE roles.[name]='ASRSysAdmins';

		EXEC sp_executeSQL @NVarCommand;
		EXEC sp_executeSQL N'DROP ROLE [ASRSysAdmins];'
	END

	IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'ASRSysAdmin' AND type = 'R')
	BEGIN
		EXECUTE sp_executesql N'CREATE ROLE [ASRSysAdmin] AUTHORIZATION [dbo];';

		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand +  'EXEC sp_addrolemember @rolename = [ASRSysAdmin], @membername = [' + gp.groupName + '];'
			FROM ASRSysGroupPermissions gp
			INNER JOIN ASRSysPermissionItems pi ON pi.itemID = gp.itemID
			WHERE pi.itemID IN (1) AND gp.permitted = 1;
		
		EXECUTE sp_executesql @NVarCommand;

		EXEC #spASRTempHardenTables 'ASRSysColours', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysColumnControlValues', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysConfig', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysControls', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysDiaryLinks', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysEmailLinks', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysEmailLinksColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysEmailLinksRecipients', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysFunctionParameters', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysFunctions', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysGroups', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysHistoryScreens', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysKeywords', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysLinkContent', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysModuleRelatedColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysModuleSetup', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOperatorParameters', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOperators', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookEvents', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookFolders', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookLinks', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookLinksColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysOutlookLinksDestinations', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysPermissionCategories', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysPictures', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysRelations', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysScreens', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysSSIHiddenGroups', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysSSIntranetLinks', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysSSIViews', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysSummaryFields', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysTables', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysTableTriggers', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysTableValidations', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysViewColumns', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysViewScreens', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'ASRSysViews', 'ASRSysAdmin', 1;
		EXEC #spASRTempHardenTables 'tbsys_MobileFormElements', 'ASRSysAdmin', 0;
		EXEC #spASRTempHardenTables 'tbsys_MobileFormLayout', 'ASRSysAdmin', 0;
		EXEC #spASRTempHardenTables 'tbsys_MobileGroupWorkflows', 'ASRSysAdmin', 0;

	END
	
	IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'ASRSysWorkflowAdmin' AND type = 'R')
	BEGIN
		EXECUTE sp_executesql N'CREATE ROLE [ASRSysWorkflowAdmin] AUTHORIZATION [dbo];';

		SET @NVarCommand = '';
		SELECT @NVarCommand = @NVarCommand +  'EXEC sp_addrolemember @rolename = [ASRSysWorkflowAdmin], @membername = [' + gp.groupName + '];'
			FROM ASRSysGroupPermissions gp
			INNER JOIN ASRSysPermissionItems pi ON pi.itemID = gp.itemID
			WHERE pi.itemID IN (1, 151, 152) AND gp.permitted = 1;
		EXECUTE sp_executesql @NVarCommand;

		EXEC #spASRTempHardenTables 'ASRSysWorkflowElementColumns', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowElementItems', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowElementItemValues', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowElements', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowElementValidations', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowInstances', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowInstanceSteps', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowInstanceValues', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowLinks', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowQueue', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowQueueColumns', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowStepDelegation', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowTriggeredLinkColumns', 'ASRSysWorkflowAdmin', 0;
		EXEC #spASRTempHardenTables 'ASRSysWorkflowTriggeredLinks', 'ASRSysWorkflowAdmin', 0;

	END



	DROP PROCEDURE #spASRTempHardenTables


/* ------------------------------------------------------- */
PRINT 'Step - Misc Updates'
/* ------------------------------------------------------- */


	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRInsertChildView2]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRInsertChildView2];
	EXECUTE sp_executesql N'CREATE PROCEDURE sp_ASRInsertChildView2 (
		@plngNewRecordID	int OUTPUT,		/* Output variable to hold the new record ID. */
		@plngTableID		int,			/* ID of the table we are creating a view for. */
		@piType		integer,			/* 0 = OR inter-table join, 1 = AND inter-table join. */
		@psRole		varchar(256))		/* Role name. */
	AS
	BEGIN
		DECLARE @lngRecordID	int,
				@iCount		int;

		DECLARE	@outputTable table (childViewId int NOT NULL);

		SELECT @lngRecordID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @plngTableID
		AND role = @psRole;

		IF @lngRecordID IS NULL
		BEGIN
			/* Insert a record in the ASRSysChildViews table. */
			INSERT INTO ASRSysChildViews2 (tableID, type, role)
			OUTPUT inserted.childViewID INTO @outputTable
			VALUES (@plngTableID, @piType, @psRole);

			/* Get the ID of the inserted record.*/
			SELECT @lngRecordID = childViewId FROM @outputTable;
		END
		ELSE
		BEGIN
			UPDATE ASRSysChildViews2 
			SET type = @piType
			WHERE tableID = @plngTableID
			AND role = @psRole;
		END

		/* Return the new record ID. */
		SET @plngNewRecordID = @lngRecordID;
	END';



	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_getaudittrail]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_getaudittrail];
		
	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spstat_getaudittrail] (
		@piAuditType	int,
		@psOrder 		varchar(MAX),
		@psFilter		varchar(MAX),
		@piTop			int)
	AS
	BEGIN

		SET NOCOUNT ON;

		DECLARE @sSQL nvarchar(MAX);

		IF @piAuditType = 1
		BEGIN
			SET @sSQL = ''SELECT {TOP} 
				a.userName AS [User], 
				a.dateTimeStamp AS [Date / Time], 
				a.tableName AS [Table], 
				a.columnName AS [Column], 
				a.oldValue AS [Old Value], 
				a.newValue AS [New Value], 
				a.recordDesc AS [Record Description],
				a.id,
				CASE WHEN c.DataType = 2 OR c.DataType = 4 THEN 1 ELSE 0 END AS IsNumeric
				FROM dbo.ASRSysAuditTrail a
				LEFT JOIN ASRSysColumns c ON c.ColumnID = a.ColumnID'';
		END
		ELSE IF @piAuditType = 2
			SET @sSQL =  ''SELECT {TOP} 
				a.userName AS [User], 
				a.dateTimeStamp AS [Date / Time],
				a.groupName AS [User Group],
				a.viewTableName AS [View / Table],
				a.columnName AS [Column], 
				a.action AS [Action],
				a.permission AS [Permission], 
				a.id
				FROM dbo.ASRSysAuditPermissions a'';
		ELSE IF @piAuditType = 3
			SET @sSQL = ''SELECT {TOP} 
				a.userName AS [User],
    			a.dateTimeStamp AS [Date / Time],
				a.groupName AS [User Group], 
				a.userLogin AS [User Login],
				a.[Action], 
				a.id
				FROM dbo.ASRSysAuditGroup a'';
		ELSE IF @piAuditType = 4
			SET @sSQL = ''SELECT {TOP} 
				a.DateTimeStamp AS [Date / Time],
				a.UserGroup AS [User Group],
				a.UserName AS [User], 
				a.ComputerName AS [Computer Name],
				a.HRProModule AS [Module],
				a.Action AS [Action], 
				a.id
				FROM dbo.ASRSysAuditAccess a'';
				
		IF LEN(@psFilter) > 0
			SET @sSQL = @sSQL + CHAR(10) + ''WHERE '' + @psFilter;

		IF LEN(@psOrder) > 0
			SET @sSQL = @sSQL + CHAR(10) + @psOrder;
				
		-- Retreive selected data
		IF LEN(@sSQL) > 0 
		BEGIN
			IF ISNULL(@piTop, 0) > 0
				SET @sSQL = REPLACE(@sSQL, ''{TOP}'', ''TOP '' + convert(varchar, @piTop));
				
			EXECUTE sp_executeSQL @sSQL;
		END

	END';


/* ------------------------------------------------------- */
PRINT 'Step - Auto Self Service Logins'
/* ------------------------------------------------------- */

	IF NOT EXISTS (SELECT c.object_id FROM sys.columns c INNER JOIN sys.objects o ON o.object_id = c.object_id WHERE c.name = 'ID' AND o.type = 'TT' AND o.name LIKE 'TT_SelfServiceType_%')
	BEGIN
		IF EXISTS(SELECT *FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGenerateSelfServiceLogins]') AND xtype = 'P')
		BEGIN
			EXECUTE sp_executesql N'DROP PROCEDURE dbo.spASRGenerateSelfServiceLogins';
			EXECUTE sp_executesql N'DROP TYPE dbo.SelfServiceType';
		END
	END

	IF NOT EXISTS (SELECT * FROM sys.types WHERE is_table_type = 1 AND name ='SelfServiceType')
		EXECUTE sp_executesql N'CREATE TYPE SelfServiceType AS TABLE (
			[ID] integer,
			[Login] nvarchar(255) NULL,
			[Email] nvarchar(255),
			[StartDate] datetime,
			[LeavingDate] datetime,
			[KnownAs] nvarchar(255),
			[SecurityGroup] nvarchar(255))';

	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGenerateSelfServiceLogins]') AND xtype = 'P')
		EXECUTE sp_executesql N'CREATE PROCEDURE dbo.spASRGenerateSelfServiceLogins(@logins AS SelfServiceType READONLY)
		AS
		BEGIN
			SET NOCOUNT ON;
		END';

	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRDeleteExpiredSelfServiceLogins]') AND xtype = 'P')
		EXECUTE sp_executesql N'CREATE PROCEDURE dbo.spASRDeleteExpiredSelfServiceLogins
		AS
		BEGIN
			SET NOCOUNT ON;
		END';



/* ------------------------------------------------------- */
PRINT 'Step - Talent Reports'
/* ------------------------------------------------------- */

	-- Insert the system permissions for 9-Box Grid Reports and new picture too
	IF NOT EXISTS(SELECT * FROM dbo.[ASRSysPermissionCategories] WHERE [categoryID] = 46)
	BEGIN
		INSERT dbo.[ASRSysPermissionCategories] ([CategoryID], [Description], [ListOrder], [CategoryKey], [picture])
			VALUES (46, 'Talent Reports', 10, 'TALENTREPORTS', 0x47494638396110001000F70000009C3100FF312929AD3939AD3939B54242B54A4AB54A4ABD5252BD5A5ABD6363BD6363C66B6BC67373C67B7BCE8484CE8C8CD69494D69C9CD6A5A5DEADADDEB5B5E7BDBDE7C6C6E7C6C6EFD6D6EFE7E7F7EFEFF7EFEFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21F9040100001D002C0000000010001000000891003B081CD84143060D04130E84A0A061048504272490600142020A1005326850A1A30206193B2848F0000204040A42462880604182021242766030A06603991B1E203000A067460D0C183C780020804F8541214CA030A1E8D181127B4A2D6A140041050DA806D8CAF56806040EA6F6DCFA940282A00F1C38C85A15A3400B061AA60D5B55C085882F0BE8956A6082C08000003B00);
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (169,46,'New',10, 'NEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (170,46,'Edit',20, 'EDIT');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (171,46,'View',30, 'VIEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (172,46,'Delete',40, 'DELETE');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (173,46,'Run',50, 'RUN');

		-- Clone existing security based on match report permissions
		DELETE FROM ASRSysGroupPermissions WHERE itemid IN (169, 170, 171, 172, 173)
		INSERT ASRSysGroupPermissions (itemID, groupName, permitted)
			SELECT DISTINCT 169, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid IN (101, 134, 139)
			UNION
			SELECT DISTINCT 170, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid IN (102, 135, 140)
			UNION
			SELECT DISTINCT 171, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid IN (103, 136, 141)
			UNION
			SELECT DISTINCT 172, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid IN (104, 137, 142)
			UNION
			SELECT DISTINCT 173, groupName, permitted FROM ASRSysGroupPermissions WHERE itemid IN (105, 138, 143)
	END

	UPDATE ASRSysPermissionCategories SET Description = 'Talent Reports' WHERE categoryID = 46

	IF NOT EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[ASRSysTalentReports]') AND xtype in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysTalentReports](
			[ID] [int] IDENTITY(1,1) NOT NULL,
			[Name] [varchar](50) NOT NULL DEFAULT (''''),
			[Description] [varchar](255) NULL,
			[BaseTableID] [int] NOT NULL,
			[BaseSelection] [int] NOT NULL,
			[BasePicklistID] [int] NULL,
			[BaseFilterID] [int] NULL,
			[BaseChildTableID] [int] NOT NULL,
			[BaseChildColumnID] [int] NOT NULL,
			[BasePreferredRatingColumnID] [int] NOT NULL,
			[BaseMinimumRatingColumnID] [int] NOT NULL,
			[MatchTableID] [int] NOT NULL,
			[MatchSelection] [int] NOT NULL,
			[MatchPicklistID] [int] NULL,
			[MatchFilterID] [int] NULL,
			[MatchChildTableID] [int] NOT NULL,
			[MatchChildColumnID] [int] NOT NULL,
			[MatchChildRatingColumnID] [int] NOT NULL,
			[MatchAgainstType] [int] NOT NULL,
			[UserName] [varchar](50) NOT NULL,
			[TimeStamp] [timestamp] NULL,
			[EmailAddrID] [int] NULL,
			[EmailSubject] [varchar](max) NULL,
			[EmailAttachmentName] [varchar](max) NULL,
			[IsLabel] [bit] NULL,
			[LabelTypeID] [int] NULL,
			[PromptStart] [int] NULL,
			[OutputFormat] [int] NULL,
			[OutputScreen] [bit] NULL,
			[OutputSave] [bit] NULL,
			[OutputFilename] [varchar](255) NULL)';
        EXEC sp_executesql N'GRANT SELECT, INSERT, UPDATE, DELETE ON [ASRSysTalentReports] TO [ASRSysGroup]';

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysTalentReportAccess](
			[GroupName] [varchar](256) NOT NULL,
			[Access] [varchar](2) NOT NULL,
			[ID] [int] NOT NULL)';
        EXEC sp_executesql N'GRANT SELECT, INSERT, UPDATE, DELETE ON [ASRSysTalentReportAccess] TO [ASRSysGroup]';

		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysTalentReportDetails](
			[ID] [int] IDENTITY(1,1) NOT NULL,
			[TalentReportID] [int] NOT NULL,
			[ColType] varchar(1) NOT NULL,
			[ColExprID] [int] NOT NULL,
			[ColSize] [int] NULL,
			[ColDecs] [int] NULL,
			[ColHeading] [varchar](255) NULL,
			[ColSequence] [int] NULL,
			[SortOrderSeq] [int] NULL,
			[SortOrderDirection] [varchar](4) NULL)';
        EXEC sp_executesql N'GRANT SELECT, INSERT, UPDATE, DELETE ON [ASRSysTalentReportDetails] TO [ASRSysGroup]';

	END

    EXEC sp_executesql N'GRANT SELECT,INSERT,UPDATE,DELETE ON [ASRSysTalentReports] TO [ASRSysGroup]'

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysTalentReports', 'U') AND name = 'OutputEmail')
		EXEC sp_executesql N'ALTER TABLE ASRSysTalentReports ADD OutputEmail bit;';

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysTalentReports', 'U') AND name = 'IncludeUnmatched')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysTalentReports ADD IncludeUnmatched bit;';
		EXEC sp_executesql N'UPDATE ASRSysTalentReports SET IncludeUnmatched = 0;';
	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysTalentReports', 'U') AND name = 'MinimumScore')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysTalentReports ADD MinimumScore int;';
		EXEC sp_executesql N'UPDATE ASRSysTalentReports SET MinimumScore = 1;';
	END

	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysTalentReports') AND name = N'PK_ASRSysTalentReports_ID')
		EXEC sp_executesql N'ALTER TABLE dbo.ASRSysTalentReports ADD CONSTRAINT	PK_ASRSysTalentReports_ID PRIMARY KEY CLUSTERED (ID) ON [PRIMARY]'

	IF NOT EXISTS(SELECT Name FROM sysindexes WHERE id = object_id(N'ASRSysTalentReportDetails') AND name = N'IDX_TalentReportID')
		EXEC sp_executesql N'CREATE NONCLUSTERED INDEX [IDX_TalentReportID] ON ASRSysTalentReportDetails ([TalentReportID])'



PRINT 'Final Step - Updating Versions'

	EXEC spsys_setsystemsetting 'database', 'version', '8.2';
	EXEC spsys_setsystemsetting 'intranet', 'minimum version', '8.2.0';
	EXEC spsys_setsystemsetting 'ssintranet', 'minimum version', '8.2.0';
	EXEC spsys_setsystemsetting 'server dll', 'minimum version', '3.4.0';
	EXEC spsys_setsystemsetting '.NET Assembly', 'minimum version', '4.2.0';
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'outlook service 2', 'minimum version', '1.0.0';
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '5.0.0';
	EXEC spsys_setsystemsetting 'system framework', 'version', '1.0.4268.21068';


insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v8.2')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v8.2 Of OpenHR'