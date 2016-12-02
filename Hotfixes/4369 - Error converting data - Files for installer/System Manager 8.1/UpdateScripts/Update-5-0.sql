/* --------------------------------------------------- */
/* Update the database from version 4.3 to version 5.0 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(max),
	@iSQLVersion numeric(3,1),
	@NVarCommand nvarchar(max),
	@sObject sysname,
	@sObjectType char(2),
	@ptrval binary(16),
	@sTableName	sysname,
	@sIndexName	sysname,
	@fPrimaryKey	bit;

DECLARE	@perstableid integer,
		@columnid integer;
	
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
SELECT @iSQLVersion = convert(numeric(3,1), convert(nvarchar(4), SERVERPROPERTY('ProductVersion')));
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

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRDefragIndexes]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRDefragIndexes];

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

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRWorkflowSubmitImmediatesAndGetSucceedingElements];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_audittable]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_audittable];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_getaudittrail]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_getaudittrail];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_scriptnewcolumn]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_scriptnewcolumn];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_setdefaultmodulesetting]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_setdefaultmodulesetting];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_setmodulesetting]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_setmodulesetting];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASRGetControlDetails]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[sp_ASRGetControlDetails];



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
			@piRecordID			integer			OUTPUT,
			@pfResult	bit OUTPUT
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
			SET @pfResult = 1;
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
					SET @pfResult = 0;
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
						SET @pfResult = 0;
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
							SET @pfResult = 0;
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
						SET @pfResult = 1;
					END
					ELSE
					BEGIN
						/* UPDATE. */
						SET @psSQL = ''UPDATE '' + @psTableName
							+ '' SET '' + @sColumnList
							+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(255), @piRecordID);
						SET @pfResult = 1;
					END
				END
			END
		
			IF @piDataAction = 2
			BEGIN
				/* DELETE. */
				SET @psSQL = ''DELETE FROM '' + @psTableName
					+ '' WHERE '' + @sIDColumnName + '' = '' + convert(varchar(255), @piRecordID);
				SET @pfResult = 1;
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


	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spASRDefragIndexes]
		(@maxfrag DECIMAL)
	AS
	BEGIN

		SET NOCOUNT ON;

		DECLARE @tablename		varchar(128),
				@sSQL			nvarchar(MAX),
				@objectid		int,
				@objectowner	varchar(255),
				@indexid		int,
				@frag			decimal,
				@indexname		char(255),
				@dbname			sysname,
				@tableid		int,
				@tableidchar	varchar(255);

		-- Checking fragmentation
		DECLARE tables CURSOR FOR
			SELECT sc.[Name]  + ''.'' + so.[Name]
			FROM sys.sysobjects so
				INNER JOIN sys.sysindexes si ON so.id = si.id
				INNER JOIN sys.schemas sc ON so.uid  = sc.schema_id
			WHERE so.type =''U'' AND si.indid < 2 AND si.rows > 0
			ORDER BY sc.name, so.[Name];

		-- Create the temporary table to hold fragmentation information
		DECLARE @fraglist TABLE (
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
		OPEN tables;

		-- Loop through all the tables in the database running dbcc showcontig on each one
		FETCH NEXT FROM tables INTO @tableidchar;

		WHILE @@FETCH_STATUS = 0
		BEGIN
		
			-- Do the showcontig of all indexes of the table
			INSERT INTO @fraglist 
				EXEC (''DBCC SHOWCONTIG ('''''' + @tableidchar + '''''') WITH FAST, TABLERESULTS, ALL_INDEXES, NO_INFOMSGS'');

			FETCH NEXT FROM tables INTO @tableidchar;
		END

		-- Close and deallocate the cursor
		CLOSE tables;
		DEALLOCATE tables;

		-- Begin Stage 2: (defrag) declare cursor for list of indexes to be defragged
		DECLARE indexes CURSOR FOR
		SELECT ObjectName, ObjectOwner = schema_name(so.uid), ObjectId, IndexName, ScanDensity
		FROM @fraglist f
		JOIN sysobjects so ON f.ObjectId=so.id
		WHERE ScanDensity <= @maxfrag
			AND INDEXPROPERTY (ObjectId, IndexName, ''IndexDepth'') > 0;

		-- Open the cursor
		OPEN indexes;

		-- Loop through the indexes
		FETCH NEXT FROM indexes	INTO @tablename, @objectowner, @objectid, @indexname, @frag;

		WHILE @@FETCH_STATUS = 0
		BEGIN
			SET QUOTED_IDENTIFIER ON;

			SET @sSQL = ''ALTER INDEX ['' +  RTRIM(@indexname) + ''] ON '' + RTRIM(@objectowner) + ''.'' + RTRIM(@tablename) + '' REBUILD'';

			EXECUTE sp_executeSQL @sSQL;

			SET QUOTED_IDENTIFIER OFF;

			FETCH NEXT FROM indexes INTO @tablename, @objectowner, @objectid, @indexname, @frag;
		END

		-- Close and deallocate the cursor
		CLOSE indexes;
		DEALLOCATE indexes;

	END'

		
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

	END';


	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spstat_setmodulesetting](
			@modulekey AS varchar(50),
			@parameterkey AS varchar(50),
			@parametervalue AS nvarchar(1000),
			@parametertype AS varchar(20))
		AS
		BEGIN
			
			IF EXISTS(SELECT [ParameterValue] FROM dbo.[ASRSysModuleSetup] WHERE [ModuleKey] = @modulekey AND [ParameterKey] = @parameterkey AND [ParameterType] = @parametertype)
				UPDATE dbo.[ASRSysModuleSetup] SET [ParameterValue] = @parametervalue  WHERE [ModuleKey] = @modulekey AND [ParameterKey] = @parameterkey AND [ParameterType] = @parametertype;
			ELSE
				INSERT dbo.[ASRSysModuleSetup]([ModuleKey], [ParameterKey], [ParameterValue], [ParameterType])
					VALUES (@modulekey, @parameterkey, @parametervalue, @parametertype);	
		END';


	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spstat_setdefaultmodulesetting](
			@modulekey AS varchar(50),
			@parameterkey AS varchar(50),
			@parametervalue AS nvarchar(1000),
			@parametertype AS varchar(20))
		AS
		BEGIN
			
			IF NOT EXISTS(SELECT [ParameterValue] FROM dbo.[ASRSysModuleSetup] WHERE [ModuleKey] = @modulekey AND [ParameterKey] = @parameterkey AND [ParameterType] = @parametertype)
				INSERT dbo.[ASRSysModuleSetup]([ModuleKey], [ParameterKey], [ParameterValue], [ParameterType])
					VALUES (@modulekey, @parameterkey, @parametervalue, @parametertype);	
		END';


	EXECUTE sp_executeSQL N'CREATE PROCEDURE dbo.spstat_scriptnewcolumn (@columnid integer OUTPUT, @tableid integer, @columnname varchar(255)
		, @datatype integer, @description varchar(255), @size integer, @decimals integer, @islocked bit, @uniquekey varchar(37))
	AS
	BEGIN

		DECLARE @ssql nvarchar(MAX),
				@tablename varchar(255),
				@datasyntax	varchar(255);

		DECLARE @spinnerMinimum integer,
			@spinnerMaximum integer,
			@spinnerIncrement integer,
			@audit bit,
			@duplicate bit,
			@defaultvalue varchar(max),
			@columntype integer,
			@mandatory bit,
			@uniquecheck bit,
			@convertcase smallint,
			@mask varchar(MAX),
			@lookupTableID integer,
			@lookupColumnID integer,
			@controltype integer,
			@alphaonly bit,
			@blankIfZero bit,
			@multiline bit,
			@alignment smallint,
			@calcExprID integer,
			@gotFocusExprID integer,
			@lostFocusExprID integer,
			@calcTrigger smallint,
			@readOnly bit,
			@statusBarMessage varchar(255),
			@errorMessage varchar(255),
			@linkTableID integer, 
			@Afdenabled bit, 
			@Afdindividual integer,
			@Afdforename integer, 
			@Afdsurname integer,
			@Afdinitial integer, 
			@Afdtelephone integer, 
			@Afdaddress integer,
			@Afdproperty integer, 
			@Afdstreet integer, 
			@Afdlocality integer, 
			@Afdtown integer, 
			@Afdcounty integer,
			@dfltValueExprID integer, 
			@linkOrderID integer, 
			@OleOnServer bit, 
			@childUniqueCheck bit,
			@LinkViewID integer, 
			@DefaultDisplayWidth integer, 
			@UniqueCheckType integer,
			@Trimming integer, 
			@Use1000Separator bit,
			@LookupFilterColumnID integer, 
			@LookupFilterValueID integer, 
			@QAddressEnabled integer, 
			@QAIndividual integer, 
			@QAAddress integer, 
			@QAProperty integer, 
			@QAStreet integer,
			@QALocality integer, 
			@QATown integer, 
			@QACounty integer, 
			@LookupFilterOperator integer, 
			@Embedded bit, 
			@OLEType integer, 
			@MaxOLESizeEnabled bit, 
			@MaxOLESize integer,
			@AutoUpdateLookupValues bit, 
			@CalculateIfEmpty bit;

		-- Can we safely create this column?
		SELECT @columnid = ISNULL(columnid,0) FROM dbo.[ASRSysColumns] WHERE tableid = @tableid AND columnname = @columnname;
		IF @columnid > 0
		BEGIN
			RETURN;
		END

		SELECT @tablename = [tablename] FROM dbo.[ASRSysTables] WHERE tableid = @tableid;
		SELECT @columnid = MAX(columnid) + 1 FROM dbo.[ASRSysColumns];
			
		SET @defaultvalue = '''';		
		SET @spinnerMinimum = 0;
		SET @spinnerMaximum = 0;
		SET @spinnerIncrement = 0;
		SET @audit = 0;
		SET @duplicate = 0;
		SET @columntype = 0;
		SET @mandatory = 0;
		SET @uniquecheck = 0;
		SET @convertcase = 0;
		SET @mask = '''';
		SET @lookupTableID = 0;
		SET	@lookupColumnID = 0;
		SET	@controltype = 0;	
		SET @alphaonly = 0;
		SET @blankIfZero = 0;
		SET @multiline = 0;
		SET @alignment = 0;
		SET @calcExprID = 0;
		SET @gotFocusExprID = 0;
		SET @lostFocusExprID = 0;
		SET @calcTrigger = 0;
		SET @readOnly = 0;
		SET @statusBarMessage = '''';
		SET @errorMessage = '''';
		SET @linkTableID = 0; 
		SET @Afdenabled = 0; 
		SET @Afdindividual = 0;
		SET @Afdforename = 0; 
		SET @Afdsurname = 0;
		SET @Afdinitial = 0; 
		SET @Afdtelephone = 0; 
		SET @Afdaddress = 0;
		SET @Afdproperty = 0; 
		SET @Afdstreet = 0; 
		SET @Afdlocality = 0; 
		SET @Afdtown = 0; 
		SET @Afdcounty = 0;
		SET @dfltValueExprID = 0; 
		SET @linkOrderID = 0; 
		SET @OleOnServer = 0; 
		SET @childUniqueCheck = 0;
		SET @LinkViewID = 0; 
		SET @DefaultDisplayWidth = 0; 
		SET @UniqueCheckType = 0;
		SET @Trimming = 0;
		SET @Use1000Separator = 0;
		SET @LookupFilterColumnID = 0; 
		SET @LookupFilterValueID = 0; 
		SET @QAddressEnabled = 0; 
		SET @QAIndividual = 0; 
		SET @QAAddress = 0; 
		SET @QAProperty = 0; 
		SET @QAStreet = 0;
		SET @QALocality = 0; 
		SET @QATown = 0; 
		SET @QACounty = 0; 
		SET @LookupFilterOperator = 0; 
		SET @Embedded = 0; 
		SET @OLEType = 0; 
		SET @MaxOLESizeEnabled = 0; 
		SET @MaxOLESize = 0;
		SET @AutoUpdateLookupValues = 0; 
		SET @CalculateIfEmpty = 0;
		

		-- Logic
		IF @datatype = -7
		BEGIN
			SET @datasyntax = ''bit'';
			SET @defaultvalue = ''FALSE'';
			SET @controltype = 1;
		END

		-- OLE
		IF @datatype = -4
			SET @controltype = 1;

		-- Photo
		IF @datatype = -3
			SET @controltype = 1024;

		-- Link
		IF @datatype = -2
		BEGIN
			SET @datasyntax = ''varchar(255)'';
			SET @controltype = 2048;
		END

		-- Working Pattern
		IF @datatype = -1
		BEGIN
			SET @datasyntax = ''varchar(14)'';
			SET @controltype = 4096;
		END
		
		-- Numeric
		IF @datatype = 2
		BEGIN
			SET @datasyntax = ''numeric('' + @size + '','' + @decimals + '')'';
			SET @defaultvalue = 0;	
			SET @controltype = 64;
		END

		-- Integers
		IF @datatype = 4
		BEGIN
			SET @datasyntax = ''integer'';
			SET @controltype = 64;
		END
		
		-- Date
		IF @datatype = 11
		BEGIN
			SET @datasyntax = ''datetime'';
			SET @controltype = 64;
		END

		-- Character
		IF @datatype = 12
		BEGIN
			SET @datasyntax = ''varchar('' + @size + '')'';
			SET @controltype = 64;
		END

		-- System objects update
		INSERT dbo.[tbsys_scriptedobjects] ([guid], [objecttype], [targetid], [ownerid], [effectivedate], [revision], [locked], [lastupdated])
			SELECT @uniquekey, 2, @columnid, ''AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE'', ''01/01/1900'',1,@islocked, GETDATE()

		-- Update base table								
		INSERT dbo.[tbsys_columns] ([columnID], [tableID], [columnType], [datatype], [defaultValue], [size], [decimals]
				, [lookupTableID], [lookupColumnID], [controltype], [spinnerMinimum], [spinnerMaximum], [spinnerIncrement], [audit]
				, [duplicate], [mandatory], [uniquecheck], [convertcase], [mask], [alphaonly], [blankIfZero], [multiline], [alignment]
				, [calcExprID], [gotFocusExprID], [lostFocusExprID], [calcTrigger], [readOnly], [statusBarMessage], [errorMessage]
				, [linkTableID], [Afdenabled], [Afdindividual], [Afdforename], [Afdsurname], [Afdinitial], [Afdtelephone], [Afdaddress]
				, [Afdproperty], [Afdstreet], [Afdlocality], [Afdtown], [Afdcounty], [dfltValueExprID], [linkOrderID], [OleOnServer]
				, [childUniqueCheck], [LinkViewID], [DefaultDisplayWidth], [ColumnName], [UniqueCheckType], [Trimming], [Use1000Separator]
				, [LookupFilterColumnID], [LookupFilterValueID], [QAddressEnabled], [QAIndividual], [QAAddress], [QAProperty], [QAStreet]
				, [QALocality], [QATown], [QACounty], [LookupFilterOperator], [Embedded], [OLEType], [MaxOLESizeEnabled], [MaxOLESize]
				, [AutoUpdateLookupValues], [CalculateIfEmpty]) 
			VALUES (@columnid, @tableid, @columntype, @datatype, @defaultvalue, @size, @decimals
				, @lookupTableID, @lookupColumnID, @controltype, @spinnerMinimum, @spinnerMaximum, @spinnerIncrement, @audit
				, @duplicate, @mandatory, @uniquecheck, @convertcase, @mask, @alphaonly, @blankIfZero, @multiline, @alignment
				, @calcExprID, @gotFocusExprID, @lostFocusExprID, @calcTrigger, @readOnly, @statusBarMessage, @errorMessage
				, @linkTableID, @Afdenabled, @Afdindividual, @Afdforename, @Afdsurname, @Afdinitial, @Afdtelephone, @Afdaddress
				, @Afdproperty, @Afdstreet, @Afdlocality, @Afdtown, @Afdcounty, @dfltValueExprID, @linkOrderID, @OleOnServer
				, @childUniqueCheck, @LinkViewID, @DefaultDisplayWidth, @ColumnName, @UniqueCheckType, @Trimming, @Use1000Separator
				, @LookupFilterColumnID, @LookupFilterValueID, @QAddressEnabled, @QAIndividual, @QAAddress, @QAProperty, @QAStreet
				, @QALocality, @QATown, @QACounty, @LookupFilterOperator, @Embedded, @OLEType, @MaxOLESizeEnabled, @MaxOLESize
				, @AutoUpdateLookupValues, @CalculateIfEmpty);

			-- Physically create this column (is regenerated by the System Manager save)	
			SET @ssql = N''ALTER TABLE dbo.tbuser_'' + @tablename + '' ADD '' + @columnname + '' '' + @datasyntax;
			EXECUTE sp_executesql @ssql;

		RETURN;

	END';


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
							@iStoredDataRecordID	OUTPUT,
							@fResult	OUTPUT;
		
						--IF LEN(@sStoredDataSQL) > 0 
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
	@psSQLLogin 		varchar(200)
	)
	AS
	BEGIN

		SET NOCOUNT ON

		/* Return parameters showing what permissions the current user has on all of the tables. */
		DECLARE @iUserGroupID	int

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
			AND o.xtype = ''v''
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
		WHERE (syscolumns.name <> ''timestamp'' AND syscolumns.name <> ''ID'')
			AND p.protectType IN (204, 205) 
			AND o.[xtype] = ''V''
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


	----------------------------------------------------------------------
	-- sp_ASRGetControlDetails Stored Procedure
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRGetControlDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].sp_ASRGetControlDetails;

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRGetControlDetails] 
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
			ISNULL(col.readOnly,0) AS [readOnly], 
			ISNULL(cont.readonly,0) AS [ScreenReadOnly],
			col.[statusBarMessage], col.[errorMessage], col.[linkTableID], col.[linkViewID],
			col.[linkOrderID], col.[Afdenabled], tab.[TableName],col.[Trimming], col.[Use1000Separator],
			col.[QAddressEnabled], col.[OLEType], col.[MaxOLESizeEnabled], col.[MaxOLESize], col.[AutoUpdateLookupValues],
			0 AS [locked]
		FROM [dbo].[ASRSysControls] cont
			LEFT OUTER JOIN [dbo].[ASRSysTables] tab ON cont.[tableID] = tab.[tableID]
			LEFT OUTER JOIN [dbo].[ASRSysColumns] col ON col.[tableID] = cont.[tableID] AND col.[columnID] = cont.[columnID]
		WHERE cont.[ScreenID] = @piScreenID
		ORDER BY cont.[PageNo], 
			cont.[ControlLevel] DESC, 
			cont.[tabIndex];
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
		
	--Fix for minimum dropdown widths
	EXEC sp_executesql N'UPDATE ASRSysWorkflowElementItems SET Width = 64 WHERE ItemType = 13 AND Width < 64;';



PRINT 'Step 4a - Remember Workflow Tab Strip'
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysWorkflowInstances', 'U') AND name = 'pageno')
		BEGIN
			EXEC sp_executesql N'ALTER TABLE dbo.ASRSysWorkflowInstances ADD pageno integer NULL;';
			EXEC sp_executesql N'UPDATE ASRSysWorkflowInstances SET pageno = 0;';
		END

	 
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

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 196) = 'Work Status'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Seconded''  WHERE TransferTypeID = 0 AND TransferFieldID = 196'
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

	IF (SELECT [Description] FROM ASRSysAccordTransferFieldDefinitions WHERE TransferTypeID = 0 AND TransferFieldID = 202) = 'Irregular Payment Pattern'
	BEGIN
		SELECT @NVarCommand = 'UPDATE ASRSysAccordTransferFieldDefinitions  SET Description = ''Irregular Payment Indicator''  WHERE TransferTypeID = 0 AND TransferFieldID = 202'
		EXEC sp_executesql @NVarCommand
	END

	SELECT @iRecCount = count(TransferFieldID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferFieldID = 204 AND TransferTypeID = 0
	IF @iRecCount = 0
	BEGIN
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (204,0,0,''Foreign Country'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (205,0,0,''Stay in UK for 6 months or more'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (206,0,0,''Stay in UK less than 6 Months'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (207,0,0,''Work both in/out UK but living abroad'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (208,0,0,''Pension paid because recently bereaved'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (209,0,0,''Annual Pension'',0,0,2,0,0)'
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
	
	--IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbsys_mobilelogins]') AND type in (N'U'))
	--BEGIN
	--	EXEC sp_executesql N'CREATE TABLE [dbo].[tbsys_mobilelogins](
	--		[userid] [integer] NOT NULL,
	--		[password] [nvarchar](max) NULL,
	--		[newpassword] [nvarchar](max) NULL);';
	--END
	
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

		DECLARE @maxid int
		
		SELECT @maxid = ISNULL(MAX(PictureID), 0) + 1 FROM dbo.ASRSysPictures	
			
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)
		SELECT @maxid, 'absHeader.png', 5, 0x89504E470D0A1A0A0000000D49484452000000B500000039080600000070216A0A000000017352474200AECE1CE90000000467414D410000B18F0BFC61050000001974455874536F6674776172650041646F626520496D616765526561647971C9653C00000C5049444154785EED9C8B75DB3A0C86EF0819212364848CD011324246E8061EC12378048FA0113C8247E8C5C7023A10459094233F2B9EC3E39804F1FC01424ADAFFFEDBC6AA1EF8F3E7CFBBCC4F99BF751EE5937996591AACB3BF93F94BE6DBAA0A6DCC360F2CF18082F75B3EF70ACC00B78B960139FC36702F09C646BBDC03804C2B291595CA7AED01B83F966BBA9DD83C50F1805662403C38040336404D35B516E34B69693BFCF42DC8254980ACF72D489B072EF600009209400F3201146006BCB4189F1733D683CA9B845832F63F95BB9DFFC73CA04006B40098B99379D50736E54FD2F48CE33F1692CDDC4B3CE0804CD5A42A539D6F7ACDAB0EA70E546FA0BE24C8FFC219010F0F7A546480BCA75ADEDB6ED1E16303F5BDA3F084F2F5AAA71AA78AFC6826A85E356C7F3F9ACE9B3E77F0805EED3BADCA54E79BB6164B4CD6DBA306EAB725FC36DA17F3805665EB93EFDE5EF4B85774E6156034763D3C369A17F380F6CABC0F1E64529D1FB62A975C2FFAA27B699CB1EDC5C2B59953F380B6187BF93C29309E1200A2FB2100F553DC341B4A57F0805ED7B41880F9E11EFC9698A8B70C15391F5BDBB1C491CF4A0B8065BE04982D0641EB7178D6186D7A777A40C14C557EFACAEC4D0EAAF4B0F5D19DC07846326D335E0ECCAE4AF33CE0C706E867046A8FCE12657ECB469B41AFF9BBE74CF056813F4EB2BF92F39F77EFC3F5F6F1803E6C15FAD248FF3DE763FD334E8DD322883F0C2A012BAD15AE63AB5EBB9F06592B7D560CD3D7BBFE2D85262D096B63FB2BBC1550E803BD02BB988556DB12B06CEDD36519813ECA7C5F43A9470435892AF3B4017A8D084F79DC04D45AA56B804E5553830CA07FAD69EAA3815A013D38876CAFED560CF8AD407D68215AF7F7047C45FBACC78A7EFD7C97F64313D85C72F7BE7E6D7FDF9BDFD541AD55A913D3D7F905CA23556AD185C46570236D80BE4206DC02D4FC655CEFB8CA2F1B1E01D49ADC766301E88F2BC47363F9F7EDC738AEE210E13EF4225AE9163F1C52F1641A60BC38649354BC79298DB1FD90CDE8B51F6F67AA3AA9FCF095A102DAFB61276B479D009C71D2EF7BE5D76CC384CECB4C3A3A3BE06F03190799E961BC77A8DEE65BE3677A6243D7B38FDA837CCE7A9DE0497C9AB69ACEAA93FD038F122FE2388E5E5BBBE9D4C15E86FFD91BE8D7BF7B05287F737624A7B63EE9A933A7FB73D50739213C074200046F39860B9483671534194F9E1B0079A48B9177BD36146280D3E2054F688A3193757ECF10C5D9AB0F8FE6CD2534F8B347A791772F96BAE984F32E08E6BEB277EA1120E7C9C845061674C9411DE91BEA243CA35B00DD009AE9985797C035B3E5B0BA66944B7CF155F3B1F0253E4BC6EC81BBE297A3EC31F381FE21B0650F402F1E3D585A44231A9C022D0002591C8D9EAC1D165B383F90837A727565E4459D84661FE841906C26606A30593BC804E4803E5DBD7CEAF79C5F2DA14AA2E1BF93898FE1CF27F2FCA8F1A4DAFB013F747D777AC2D3EB59F223E76C9CE4875922B1267342570298DA91A995E49B5FCD77E88EAC712C026C8B581D912BC2F7B3EB8F861281ACB5AEFB56D6C2170399183F31D4C92C5598453A6541C9CD39CA4202AD02A299AC4A97032B012A1F05DFEDBC3C4F2FEB878C7EA68BECE7498D2F46FD0BF2A1876F0E6AECB6D1E241E2F9F1AB20C7F3833681391A9E590BA75DFB3841E6BEE0705B1A7B3A59A00294C608FC2098A70AFF925372C7D9F112A8A3849955376112B51EF0FFDDE5B03258F1A11FC52066345F8D40E73763C94FBB8C676F128E4927E7F3C4A802109DE5CCC1C99DF4FC0038D3A9E9D755412DCC70DC50011C5BA3330B0EF047674E5707E4C1F1678A06171C530335803A07364C022434FB80AEE9F82059B18D20E681EC01750F78AABE92CD9323B8E8F5AA9CF7856AE8496C39E36FA6DA730E71796BF15C0DD4C228EF8F4AF19E55E04A12149DAA72A2BC291ABC04D49A381158F32A12817FAC5C0178A96604FF50B1DFDB782B507B99DF2DF004B661938DA302D65AC1E8D3FB7B8211390F0F1B5D89B60AA885C9CE311A2A819ABD4ED2E07683549D54A20FAB821077B71F0AEA88DE3F0F44AD47E878D1A3D59A457EB83AA80B3E6A56FE00D41E84913DD575CF5708CF8EB8EB06FC11A8354807C7848CABB50727D9C7683F498268CCAA4505D4B3FED89CB314D40A6C742D8DD416C986B7DBD3456D13808E6C651D9FE03FAB669EE7BF0C6AEF87EB825A247185FA20A52A2C6BD1C35F05BBE1D6ACFADE10D4911D802F0267F880ABA0F586C2A39800EAC7570035496ACF08DD9F59A5BE0DA8450AD5F8ECA4F9371AA74BD05B3933E94F2BA0AEBD7B5DD47E28A8F2A7F81C9025958BAF22B50078FA666F9831BF45A5CE6FD830E16ABD7696BCB37673699F9EE1EC3A95BA00E83140BAB732A6A7EFAC85390FA4D19824C04FDA8F468B11256EAFFC8F56706F0DEAC2ED50FD5D41A4BFE84D75B611169A96FD2E7647C72F6C31A3EADE9423CC79383A3B2183FCFCE614F006AD05EE89631A8953AC0C9544A83AA99140B97DDD3D7DD3D17FDB383FAE5EA915D41E401701B2E0B3AF1E7B2B49B2CB7C31E2AD72663C52955D50966A35119001DEEBD2CCFA06802657A1D0223B1A0486E4A3E5A0AA1F6AB425A3B14B26CE44CEB972DE6F85C1535D3C6DF56A2FF8E256A0CE6FC1E6752FBAD2B68C74EA3BEF337E6E0251932AC52CABBAF95BA6AA4E727EF22C1482BA101414FDC884D7DA82096D0548117EF2F7C36B3D8CCEAA2B019239987315DC2D5C57032787F3FE1CFEEFB91FD4CFC782B05B819A643E65F28B20529BF64A9BFFC2041FFA817F6A490F70CDEE99BC824EDFBDBE2B821AF0CA44293F660ACAE621887CF73556E181FCFC561802794B96C760E074756C5E29F207A812FFE60351601B81B479AA287E1350BB6A99AB82FFF73201EB4E66EEFB5271803E1FD878503EF0C27678FB5102755EADA1E71CE74B3C467E3350CB4E09D02503C8F068345B0F132C0C6AD5BE04B6DC21251DA0D907CAE11064B21FB604B23754EC63ABF7266AF13131BB4CDECD40ADC0CE9F9D1AE697FFAB898ADF6BFCA29B218A61CE6B423701B55002D452104AD766AD1D6806DC81BA961CB3D760AAE3A1E2210CE4EAA7572B8D13FBC52BCA2D0A4D2DD996DC44D877ACE8CB5E02F03D41ADF2EDB9E25CD1177C846D85ABFCB518C11E19C42A2C2CCA0B9C45FA10CB74DEEB9B83BAA448D72B951648D6DE57E0023CAE342695A609D6B5F5E8E5A7FA12A067D1971B9BC2E0E75BAFBDAE68D979ECC67EBE7717BD8C8FF90E3E6D1E0A8A52827E2F3564A3DF3CF0101E103453CA4BA39D110F61C1F80F50AB57E383A8BAA9716D0F0892C3FEF1DAB2D7E0CF55A4367CAEC1EFD178A86D47F964EEBBAE5E3542CF763FB82FB55DF8D387A393E91716152A66C4DFDA1CF657D159981C822A1D2AB1D4F86BD0ABF1A94FBB06FF47E129F6D14B0E1AF81D00D2E093CC6FFAF3E8030388022E25BCD19010BA9FCEB9757FDE1EB43F741FE0A6DE38F789620750B3CF73CDEFEC4CE2A16B094FAA83E99D6C90090FE3E37536D95E17F62776288FBF3AEA9708D3ACBF3F4A70D52138DC1E181E4AB76BF949ED3550A7E0AB2F8E063402E5D6583FE839026D4900D949677AB3C45E36F12F0339D0721ED9F0E33381D60195F3D08EE055BED0B2078F7453C0B4A03734C8808E69DFF904D067E5CF1E71879661F4C84667E80E3207041991D2CE3E26465C2B7035BEAAF4B71AF785B1F7D0E35E323598048D400F1A4002C9F7CF0C30D0EDDCFA04D44A9BD66452EDA0B72241F00D60F839150D45043413E0EA1E3C0E0E3506BCC1152103730DD449874C3F626EC9879C93EA67BC733BD0F9A307D4189D8CBBE550876314CEDF971C7A4B7DEE294B41E77F138A4F2CA039A8093EFE020080AD066AF688AF81DA4095C0A44085E697E319163995C5B1516696704B418D5E762B91C413DE996D001A9DCF3DA086D1207396A56B065AF8BF65CE43C1AF35653C2B2F05DD493F770A44ABD47C27F80618FC0608593B6681371A4B087C7E565A032E4961D516998065AF7CE03DF9F300F9CE1A3AC0939F07990640CEB237B63F0A72F4621D39C8E72C3F739E33A61FBAB06FBAB1EF93D4D3990EA9FD3005E4C7EA80398257B9FA55AE198D231866D42A329E15C4B9DE2EE8F89F990A0C9FEA33029A5759D6882D33150747E3D7E0012DBEFF26BE2A237DD773ACF31DBA496C9C6E473D97F6759D337B74C8E47BBDE18D3E2617FBBC7E007A949DED253A77161D92CE28C09725C30C205392C2D15081D01980397B56617CA27052EC5540B8D9F1001E1040913906B425E0BE8476904364E3E70398BEA9F0CA1ED0B2BEB462B7404DA28C55FD95FDB7D9F6C01ED0AABD934F2A6AEF00B84CCEA52ABCB5130F1CE41757ED7F8A33332B0248AB2A0000000049454E44AE426082;
		UPDATE dbo.tbsys_mobileformlayout SET HeaderLogoID = @maxid;
		UPDATE dbo.tbsys_mobileformlayout SET 
		HeaderLogoWidth = 175, HeaderLogoHeight = 47, 
		HeaderLogoHorizontalOffset = 5, HeaderLogoVerticalOffset = 5,
		HeaderLogoHorizontalOffsetBehaviour = 1, HeaderLogoVerticalOffsetBehaviour = 0;
		
		SET @maxid = @maxid + 1;
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)  
		SELECT @maxid, 'absSignIn.png', 5, 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D0000002574455874646174653A63726561746500323031302D30332D32305431353A30323A34322D30353A30307B94C7050000002574455874646174653A6D6F6469667900323030392D31322D30315431303A34313A30302D30353A3030A21409A40000001974455874536F6674776172650041646F626520496D616765526561647971C9653C0000088C494441545847B597095454D719C7276D726252B5DA92D8B4E25A3D8A0717824663D5065136990111641986615F64976518242C2A626513354454B619766407657123624550418D82B82088803B1E0165C9BFDF7D2964ACC9396DA4EF9CDF61EEBBF7BEEF77BFBB1D7815B7F27823F41784F2FE5714FBFF9ADF5CF0E9D3A7734C9D3A9563CA94299F6968689818181824585858D42B2B2BF7110FA82E9E10117F1E69FB6B822AF6F94F814FA6292B0BFF3A6B66DB8C1933A0A6A6069280A5A5250C0D0DA1AAAA8A69D3A681829F23663289B1141847A394CD9C391373E6CCC1CA952B21140AE1E2E202535353AE3C6FDE3C6869697152D4AE8A043E1A2B81F76864E3088FCFD5960CABABAB63D5AA55D0D4D4E4FE2E59B2042C236BD7AE854422814824C2A2458B582694C64A8047C17994015E6E90F54042A8372283FD917EF81B5416E7A33053862F977F0147474788C5622C5BB60C0B162C60027F1C7381BB89120C5E3D8EC18EEB187EDE85E1178F30F4B81D8547E2A0A7A7071D1D1D6E2AE6CE9DFBFF11684D0AC0C0A5020CDDBD88E1477730FCA41D431D3770A722037A1AAB41BB032A2A2A9C004DCB9431CF0013785D23C3E0B50A0CDDAEC550EB450C3655A3A72A11C162017434D6406DF1422C5BACFA94A660B295AD3DEF5D18DD86236B8009BCAADA8FD767533050978B81FA3CBC3E97869747C3717F9F0BF23DF5A1BD6209ECF96B1B49E0FD7709CEFABE25708F0994EEC6ABF2184E84D15F16859E242F3C8BB5C276A3956C0B22DA69D3D7EC1C187381565A84FD85DBD15FBC13FD25BB385E6649F16CBF2D9E469923DB430067FEDF7BED36AC9E3D66028A77414B82DF33BA0F4822EC470A42F0E2881B9E469AE3599C15AEEEDD82F65469D0C89D6165EB40EBC0C19E4824EA156065F69ED553A67E9E378E52F6D1C638EFDC17B9C114982488DEAC003CDF67CD8DBE25D61E9D69814059B8366B2BB6737475F3DCDA50585CD2D1D4DCFC921E8CC0CAEC3DAB67ED08D6FE2D380161F23A0EF6D1D311CE16D7E3BDD19B138897E97EE889B7E782B7455BE34C84137E280D476EAC8BB7D8DE719B3C3DB3F5C99327432CE883079DB874B90155274EE2E2A54BE8E8E8E064587D7A66D63D6B7B2709A14CF014794BE058A89DD21137A3826B917683DD7136781C6D899BBB4538B1C30EBD79A16CF4080DB2467159493F0B70E7CE5D1CCD2F40ECC138A414A6425E2E477CF6B7D8BD6F0F32B3B34199E0442AAB4E74D938384B089E22A302BE15222E03EB03D4C77DA634B1AB7A8FD3504E8025B2FC2CD0B8DF83A6238C0B5EB95F8ABC4C39F7D193A74F63574C240ACF15A2F27A390A1AF2905E27C3A19A6FB1F74C1476E5EF4048C40E549D3CC5B54FCBC86AB5757471257823BC21A0255DBA9CC89AA43CFE87B64C09175091BEA21DD8131AC0D28AC62B57B133660FB2CF6722AD5606F9F954C49C8C4478792882CAFCE153EC0ED7420738E5DA203022849B9EC7341D9EBEFE0D76CE5B26123CC6A8008DFC430A7E6AF3C1AFA0663C174243755C4CF67C43A0E660100AF372D0D3D38314591AF616EF45566D066A9ACEA2A5BD05ED0FDAF07DEBF708CF8F8065B6314CB30D609CA3076B9910492932AE5F7E61D17D7B173731C1638C0A68FA7FBEC620F26FC36EC5C6B08AD78786CB527C326B12262B4F20C663D2D4F1D0E2EBA2E14A236EDDBE8DB0987048F3FDE095B8150EF14E308F15A2E872013A3A3B70B3AD19DE726F181FDD80CD051B6056A20F6954105A5A6EE1465353AFE3167719C1638C0AACF5518BB5926B81D602BC4ACC609F2180BAE97C28CDFE3D267CFA31FE307D02CCEC2CD1D9D5895367AAE1BBDF0F9B138D2048D081E0B0360C927560106D88E68E263C7CFC104D6D4DD85CA40FB3527D5894F3E195E48113B416BABABB879DDC3C6A081E433103D71D0BF970A50C781C33816BA9111C0BF421CED28628753DCC1234E12275457777374A8F1D87538C33F809DA102452F0141D6C94EB6263DC463C7AF610CF7B9EE3D6835B303FC68745051F962705F0C87041514929D7DF63ABDF256777AF0F899F0468FE878587D6432CD7817389000E451B6097AF07DB1C5D8833B5C1B2E314E884B6B6366E55BBC5BA4390A403C3541A7988095A1FDF45CFCBE7E8EBEFE3E8EDA34389A86FAD83799A09B6CA3D515E5985F6F6F6E12D9EDE35048F319A81D55B16B66978ABC1F4A006ECF22870AE2E6CB27546838B64EB611F6C83FACBF5B872F52AB64505C350A60BA30C5D6CCAD68328CA162FFA7B303038805703AF381EBD7808F3741358FFD310C10742B97E0D8D577A5DBD7CD2091E8313A0E737F3B5A65B2C14CCEE591FB0F4A7C0695ADCC8597061CA3A98871822293305ADADAD48A65D609324E256B949FE069816EB63D33FCCD03FD087FEC13EDCE969814506053F6F08D70A1B24D22E60FDB2728FDE77F3F615133CC688C07B24A1F4175525DBB95F4DAD5C6EA3F25AFBEB65D814B7860BCCB048D4848E74053CB679A3A9A909676B6AB03D9AB65BEE266E959B1FE74358258049B439AA3BAB10716A07AC6B37C2B1C60CDBE376A1FABBB3686E6E1E9604853478F8F84F24788CD12920810F88291F8C7B7F056D3D4765B54FBF99B76EDA77AAFC591D5F88E6F72F17AB60217F3674445A3870381E376FDE447169197646ED815B8613842704109D3680D5594398C699C3A66E233CCB9CB8C38AB563ED53E4E9F73C7D25EE046F8437EE023615C478623AB194D0268C0921614588190223E39CCC9CDC9E1B376EE07C6D2D32B27310B68FCE8524297C33BD204D9522EC403864E99994A97360EDE800EAF2F20B90103C4546A680AD03C587897C4CFC8998432C2696132B892F5516A8F22DC436070E2526773634340C5FBB760D17EAEAD8858382A26254D06AAFBD7001EC3DAB4F4E95DFF3F6974A086582A7C82F0928CAFC960AE3FE9D19969DDF111F111F521ABD0343C21A33B2B23BCE5457F736363662045666EFA9BE61AB24D09DE0FD1CFF8DC0880CCB0A5BAC6F3C3E01DB2613F6848CA8578095D97B56CFFB25DEF99F4B1F6910EF5DF81700D039794502AF3A0000000049454E44AE426082;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 12;

		SET @maxid = @maxid + 1;
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)  
		SELECT @maxid, 'absForgot.png', 5, 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D00000025744558746372656174652D6461746500323030392D31312D32335431313A35383A31362D30353A30300FCF81FB00000025744558746D6F646966792D6461746500323030392D30382D32305431333A30353A31342D30353A3030B1F866360000001974455874536F6674776172650041646F626520496D616765526561647971C9653C00000A07494441545847C557F75793D916CD1F32CE9FF09E6D1C9511CB80BE5154060B4A110C9010020408D211463A842A5D21849204A5457A07E9100DBD0B4355110512C536BADFBD979779645CF3F37C6BDDF515C8DDFBECB3EFB9E7720070FEC9C1C0D56A35A7A7A787D3D9D9C9696969E1343636726A6B6B399595959CF2F2724E696929E796B7D75E0F9148E8E9E921157B79AABDC5E275B1D8F38B87C87DDDD5D5452D10384B9D9DF9428180BF572693710A0B0B394545451C8542C1292929E1A8542A4E5555159BBBBDBDFDCFA0D9437777F7DFAAE0EBEBF33301BD1F1A1AB2259349D1DAD28CB1B1512C2F2FE3D5AB57585A5AC2F0F0309A9A1A71FF5E36FCFD7CB61C1DB9393C9E83C9DF29DBD0D06048A0ADADED1B02FEFEBEDFF9F87845C5C6C668ABABAB08E012B6B7B7D978F7EE1D743A1DB45A2D1B6FDFBE65EF74CCCFCFA3BCBC0CB76F076B1DB876514E4E37BFFB2B11AA84FE1B7BD8CD28252585E3EBEB7DC0CBD3A3BCA0A0007373730CF0CD9BD71818E8476D6D0D140A393232D2702F2B03E969A990CB0B51535385DEDE6EACADBD64A4262727919D9D097B3B9B7207EE8D1FE8BC7AD0B2B2324302D5D5D57F7E108B3D0EDCBA256EAD7CA462D1BE7FFF1E1D1D8F91929284E4E424142A2BA05475A0B47E0C152D7378583B8AE2CA4EE42BCA2091C4212121162D244D9B9B9BD8D8D88052A9809323B7FDA69DCD413D81E2E2624302151515EC83A7A7688F9BAB8BEA1101A7C02F9E3FC7C3074A32693CE40FEB50D7B90065DD0C52952388CF1F428C5483A81C0DA2A543482B9E80B4621A195215A22223904FFC42D57BFDFA3551AC08B636D7547637ACBEA73872B9DC90007529FD83502888C9CFCF63395D5D5D812C4F8AACEC1C94D56950D13A8F8C076388C9D52028B51FA2D82EB84577C235AA1BC2E86E7848FA1194AE4192620269457D484C4A45467A2A666766F0F2E54B92AA14DCB0BD164B71F2F3F30D0928954A02EE6C127EE737DD0CF90125A054CA91997D1F95AD9390D7CCE0AE62840173435AC10F6B473421A2AC7B0665FD1C240563E045F6C23EB4075E491A44E48E23A5680871F149B87F3F0B2F5EBC80E6E913F8F98975376CAE99E6E5E51912A06B95CF73CCADAC7CC4C01FB7B7B17C96370C32C9D394A3B893A5C675BF46A4C84730BBB889AF5FBFE2F3E7CFF8F0E1033E7DFA84D9252DA2F326601B42D449184478EE241265BD08090E4665A50A2B2B2B901715C0DADA529A9393634840E0ECB43730C057BBB0B080F5F575927309648A6A54772C30D91348BEB9B75B117EEF093EFFF1051F3F7E6444A9DBA9D9A8C4CF895F3636DF22BD741EDC700D7CD3C72091CF2136F521424202313D3585C1C141B8BBBB686D6C2CF71B2C43679E836B0E918A4EDADDDD85B8D81834F42C12534D32C385A40FC0E94E3B56D6DEB168E9B2A4D1CFAFE8A099DAC0DCF2CEFA9F9E9EC6CBD7DBF04A9E80387502213944BD9205F8FBFB4155518ED9D959C44B62606375D5DD80009FEF206D6CA86785849AF0BE548E1A127DE6C3B19DDCA70D20A96898C94EA37FB3B98D84C20970C3FA7135A00F0E51835826E4C6C6C698EB131F2C821F3701AFBB5390281611119785C4F838561B14A46658595E92191010BAF0D54FD46A26675858288A4A5A515C3F8B0C4220AD7894986F00C9F25112A996E53A2277143F0BDB7046D48573E27E5CF07D8A91391D464646C8EA5945D6A3555C09198765E8381CE3A61120A984B7A73BC6C7C7515F5F876B572D9E18107073755E9F9C9C60EC852E3C3C6A1E477EE534EE954D2082AC73DBA0769CE2D7C2885B0B635E13016FC7198F6E98790FC0DC5F0359DD732C2C2E9202D4422AE11A62142BB81E3643C634BBDF08ED2375C012A3A3A3A45AF6E0EA958BEB860484822FB468D0CDC5DAEA0ABA0657915B3105FFBB03B8E0D18423F6D5F8C9A11EC7F9CD30717D8CFF78F6C0EC961AE60183286D5BC3327138D94FC8E4BD985CDC062F6101B6D1F3B08D9A874DD41CECA2E760617E16434343D06834B87CE9FC1703022E02A7F529E252BA029C1CED98026E31BD38C9ABC3D19BB538E6D88093825698BA75E017AF5E9CF779822BC123286E79C576C58E8E0E747575B1F4F8643FC74DC9D2CE885B823D19DC480DACAE5FC6D3A74FF1F8F1635CB6386FA800CF89ABEEEFEF231BCE1B04F8DF4242763591BB868C3A183B35E214013FEDBE93EF8B7E1A5C0A1AC1B5DFA6F0F1D3571635E92530B3B48580DC3538263E8753D20B36E8332F650DC2B04AB80A9DC966360095AA02972CCEAB0D14E0DA594BAB48FDA726CCCC4845747C9641E4BBC12F078FE2DA9D29D844CE9139C0A2A7CB3258B60E3E0173BEFB6A67A4EEBCBB656DC12B2809E1C4DC7D7D7DB84776480B7333C35540F2EE9A9C9CC00A4B33692C48530163C75A1C2786A339A7B25FF07DF267E43691CF6017BBC008E82F0A28487B0D1736DEC025FD0D84E91B10656BE1221482363394C0ED607F42E09C611DB0BC62BECF85EFA07DF66C96D5EDC0001F0445E7E3944B1B91BE13E788DB7F0D18C2D5D0095847CCC22E6601DCF81526339F444B8129986BE626DC32B758D4F42E967E80F0760144EE2E206D18EAEAEA607FC34AFBABF9D9030629A0BBD305B333D2A2A27C6C6D6DA1BAEA11480F086BBF6A9C15F7B1E8A9F474595177DF942CB3FCEEBE28A03B8956744F47EE3A062E900C43201030D9A94FE2E3636176CE248FF68A0604E8EE646676C654E0ECA81B1919665E48BD9B84D0DFC271C9A71D1681C3ACB0584750E917E190B00A5EF24B864FDD4FAF1DF0B78C8077DE47B867AC42240E44D89D10065E414AF14D3B2B9DD95913D36F1A12DAC55246674C4FC6C44447B04E86EEE3F43928240CF6C12DCC74347AFBB845163D35DAEE8B46ED99B30DBFA2AF70891F833B010F0AF4450329F1A4E965ABEBB4A931EB07BE69C9FE4FE0C49E93278C54599969AC26D0EA9842CC493A25788717C09E188F39FD7F8613666CB0C8BD72DFC3A7E00FF8E47F22392F8233919D464ECB2E8D3E32E20E7E3E715475DAC4987544B445FFC603FA0F278F1F3978DAE4787B7A6A32ABEB741F2F79584C3CE1010F0F117C4353E1165103BE641222623641E22C5C231BE0157C1742A12B3C4442D68C52601A797878284C4C8CDB4F9D387A488FB1BB07654CE82142FF47DABDFE6474F0A0D1E1031521C10124C79DF8FDF7DFD94E5752F2004989F1B81D1400673E976C2ABF82E7688F4022AF24361A8585F9CCED14BCBCAC14B7C422FC74F480EAD8B143877677C5F5F5F5860AEC26A02772E4C77D7B0E1FDA176D6569A1A369E82644E86E4637147D4D27272AF4F7F7B3F54D41E9A067082A39D970744647F6471F333AB8473FA7FEDEDCDC6C48607797FAD77F3E7CE8DF26877FF8578EC5C5B35A7F5F31595219A02D3BADE914984CC63AE7E4A4787879B8C2FCC22FDAA387F7E5181DDE67FAD7B9F4EFE4B7DF36A5940455821C4658D7AA1FF49D7E2739DC4FD6B0DBD5CBE6327BDBEB6A22FF3AA91D5F4E9B1C5B3F697C444DA2951DF971AFDBD11FF7EDA767497AA6D49F076B6A6A3854F6A6A626762EDC7D14FC474FC654917F9CC07F01656C33725C409F460000000049454E44AE426082;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 13;

		SET @maxid = @maxid + 1;
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)  
		SELECT @maxid, 'absRegister.png', 5, 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D0000002574455874646174653A63726561746500323031302D30332D32305431353A30373A32312D30353A30306A3A1F5B0000002574455874646174653A6D6F6469667900323030392D30392D31385431323A31313A34322D30353A303035876A890000001974455874536F6674776172650041646F626520496D616765526561647971C9653C00000937494441545847B5970934D5DB1EC73751CA8D0651BD10C958A23CBC8B22538EE9C8980AC9748946A741A5B884924C458306524A8464AE542FB1D2AD9E1B85BA7845B712329EE39CEFDDE7DC6EEBBE56EBAD77DFD26FADEFDA7BFDFFFBBFF7E7F7DDBFFFFE9F4300903973E608242323335D5757D78EC964A6B8BBBBDF575252EAA5D7DFCACACA66D0D6974AFE8FB1FCE7C6428249E8A4DF51D969AACE7F2A2F2F0F4D4D4DD8DADAC2C3C3038E8E8ED0D2D2829C9C1CE8983AAA797C88B1589C3F8760229A6198A2A22268C6D0D3D3C3AA55AB10181808EA020C0D0DA1A6A6062B2B2B383B3B838EABA60093C6148066676A62A037A2ADAD2D0030323282A18181A03F77EE5C98989880C5620980F86E5000993105A00E90CAE840CED5C4DD381EB70F05D967507BB302F7AB4AC1B03083B7B737BCBCBCA0A3A3037575F5510A203DE6004D69A13CF6C3AB18ED7802EE87D7E0F6BDC5E8BB0ED41665C3D2DC0C0C0603AAAAAA505156E67E23001677E4FE45705A6AC0ED6A06F75D3B465F35E1CDDD7C385A1AC3D4D4141A1A1A505551E6D26D99FD0D1C60F146AA4F82FD531138CFEF60B4B5169CC65B18BC7906C737AE822B6339F49768C1446FF1207560CA98033C4B6761B82C0123B74E60A4261BECDA1C8CDC3D87FE2B51E84E0B44C5367B30FEA185FD9E769D1440F89B000C154563F8FA410A7204C3E589182A8E435FC62674277820CAD9000A0A0A38BB6DCDD16F710E90C663A1DCA182080C15FE88A1A228DA4661E0E22E7C48F64677BC3BEA22D620D6873914B69A3177CC0186AEEE270DA9DBDEF7E7ED03ED430092BF8F661F8CEE43EE14621D6A6303D07E6667147F2C5F63BA05FC091F1CD974A1332BEC7700AA810BDBF121C95390FDC3586FBCC86001D7A36DC61C40B2A845905155548066754C00B73B33141F4F87A027C90BDD8757E35F315E28D8EB055E7134228E45A7935D718AFCF1FCE7BEA6BFEA0C59FE8C2700A88CF217CEDEE61E9ECD721FAC8EF54755B41F72586B5016B11EC30591FCEC71F142127402F7720CE27246197B53DA1D43234B3CB7ED39171CBA23D8C56FC372BBF5413AAEEB03E62C77F514A122FF8B3E0390CCAB334966C14EA2B1A88BBFD8D7C4A62E2CD9990CD6E33E6CCDBD05CBF33F41EF4C3DD40FDEE4AAEFCC629B459E1B72D997F6C67777F47DE7CD7B73FDB66C8FF30E0E7170F50F5AE2ECFB833C059AF02594008064155A0B5D2CED9874E91E261ADB60B038E6AB00171342B120280AB79B5E82C71945EFFBB768BC5385B7BDBDA86F6CC7FC843ACC486BC1ACD4E7908FAEC39C2D9731D72F7554DB278A6DECB3ABDFCE6FD36307FF0D45CE3EBE69EB36063AC79D3A2A4EC8D97C21925D7C6D66592B96943743CAD41E7B42FDF06B611CDE17C5E27D610CDAF20EE0607830648D2DA11E1883F24E3678DC51FA615285AC8C14349495C0E5B01192530DE99456A8EEC8C1545629C647B462FC963A8CB3AA80884525C48CF2F19DD9694C631C858269189639F9FFC207583231F7CE8061C3080C6BDE6049C12368AF0FC1DAB56B111010007F7F7F38797A63B6E7464826164035F010AEBCA200EC01BCBD128BFEEED7109F381E3DEF5E63D589060847F68030A340E456608AEF035C6AECC681B22604E43E47D0D90628EEB80BE2560251871A6831377F20C269D9E1B2E5CD3CA39FD930686043B7EE3D16A7E4C2797714DCB785C185150E93036998117F0992276E42C52F1699AD230067103D3B8C805F2B71D26705D823A3308FBA0EA21E00E21207B23613A26EF730F3874A4C762DC25CEF12845F78861511B520BA2741961543D731A48D88249CBEA956F9128B6FBF86CEC37E683DE883C6FD0F985FDD05B98A76CCBCD602A982264CCD7F0AC9BCA750F18DC3E9A68F008FD6C0A685E01CB207F769057DFB00ABF02A10BB5C109BBD208B43A016D68CC4FB6F107BB515AEC9F5500C2AC522DF0A885B654244AF002A66415D442422B95121A30ACA397550AFE986CA3FDF61DEED3790BBF11AB32A3A30A3F4174CBFFE1253AFB5620A95B24F1CCA9EFD1B1CBAE0B50BA9E0A4EA2276BD156EB775C33BB292029C0731F00131DC8E091E35D00DADC2429ABD4D6C1DDC0ED7A2A2FE2DD26FBC40DEE35EB86C8EE822421E8129930276F36627174281662D7FB31373AA5E61567907A44BDB048B4F395F0789133730A5A805F3D7C743615F39A44F3F827842295A8E984132A71242CF3E40645D3C48C04310C75410A37D900F6D445A491B96EDA4CE689C02D13A0F61AD0C88689FC4148B1BD077D8D249687C4F8CAD2F8BB804F44D4B2B854CD173C894B56146C9A7CC8B5F4022B504930FE542B2B019DEB169F8B17E082A9B6F41C2B51464510688453E2D3CBAEF064920AB6F816C7F09B2865E338A01B12A84B0CD390A7595BA930562461DB2B80C327D37165A85080064A94C88ECBC4D64A94DB190F59A96F15EA1BD13B6C66362E45988D3CA17DF910C31FF704C8C3E8F98F474C17EF3831E05E81E1EC593C71F117EE2299436D460AA712E44969681481880E85DA015DF00624A1777790CA2130CB2D017E4FB08109534683BB0BAF800225452540B89E4347BA2B4308C1858E50959BAFD2C64E3D12964E7D527BCCC6E8018590F0859AFFD181293C41D66F33E437CD9E1D25B9DDD7D484B8E87DBCE424868054151E708C64B6D05119D09221608224D9D12CFC0020B5A847F8AF19F409469AB4F65466543C5FC24FB71E3C63135D6C50E491F1984E5F12EA4DE7A8DC68E7E0CB0A9157F8A47F74AE861E68594980D38BCDF0DBD3D1D483D158D2BF977915ED403E375059836EF0C14F4FDFF03E00F9671B4438F48329DEA6F54729FB66936FDEBA6F87733EB2B2A8E1BDB66B8A58C48B85570C59C1B30C9BD199A9B9FC13DE505EA29D095AC44B4BE7C84FAFA0270E809F97B703F2372793CB0393CEC494E12D4C07F0B217A932F41585A5A8E33633ACD5FEEE0B2D2D8DE79EB5286FD117D8643E90286D7D399E67B7A24969F1F16D32D87E8EC30287AB543DF782B9272DAD0DEF10EBD0323F872E3F6A6A474FDA55F361440F089A5AD1095A8B98DFD0C73BB952A66F64EE6146AAB31D329416F856D9EB605F3B6A27568FB74A39491492A99A3A2D32F414C3A0BAAC67970082E4362CE13F4F6B311919EDAFE7F017CF949A530C254625453CD6D996A664C672B133B277FEAD27E235BFBA38B573856A8587837492D8B1C9CBCE0185B54EA2CC4E4AA21A71FC2191380AF00895018712A290B6B5B0AE4C430B15DE9696CE7146A68C38CD7B1B4C9D2B4F26C563364E4FC062F51754739E2DB4E0000000049454E44AE426082;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 14;

		SET @maxid = @maxid + 1;
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)  
		SELECT @maxid, 'absToDo.png', 5, 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D0000002574455874646174653A63726561746500323031302D30332D32305431353A30323A34312D30353A30304A7CDD980000002574455874646174653A6D6F6469667900323030392D31322D30315431303A33393A35382D30353A3030E3F71CD00000001974455874536F6674776172650041646F626520496D616765526561647971C9653C000008B8494441545847A597775454571EC7DF0C43114551572CBB6239ABD9AC12546CAC468D6635B16413F7240169BA2A96985816A40445948E0C52041D40A40D10C5A10E75600AD54141CA309401440404141122E8AADFBDF3804462C073C21FDF735F99F7BE9FFB2BF7CEA30050A349E0B99C12FAACF5C8BBB2AD5512B0519ECB5E2D14787EE49BEDA177248E1B991D11CD6B0B088AC8F6B87859B3BBBB9B52AAABAB8B22F747682C8F51CD950F15857E256E957E8BBE86FDE89699A35DFA0D14193B5014678C1BB73210C1E5C1C72F04A1EE07BB88A990884D7480C8702CD3B7EF8D0990CFD981971DDFE145DB110C3C3C8CA78D47D1D764899C743F845E8F83DFE550B8BA5F4214C70ACDC27F419EB815A5511B5070650D08C46B22DEFB40C60490047D8A979DC79099930381480017371FB87906C0FBD265B87BFAE0BCB33BECEC9DE077D106DDE5267842D4556A828E923D682D3042EE457D2548C21F4E81E4F266FC8F00084462743CEE063F4300B66F20EC1C5C6065731627ADEC71FC941D8E9F3C0D3FB62D0AF9D668293A80B642233C101BA1EDCE3E2500F11FBDCEC68C8098002853902B96E0D9F3FE11AA90C9C14BE4E37260301C1C2FE0D0D1E3387CF404A2A22320CE894553EEB7682ED83B4E8080CD78D17E1442491E7A9F0F0C6910A497D6D0B5FE01B4767422353D9B1E7344122804466810598C0F4012F0095D7C4A803E62F22BC45BC7CAEBC3A281FA914322A6C836426DEEB80136A1FF81254479F93480ADAD2DBCBDD9080C0C02373616297C3E84C4AC42568D1A85023D7DCF69C80C412EEAB38C50231827802860139E371DF805E0C9B35EDAA8ACA282188B9194924A83F80790CE2060D6D6D678DCD38BC4B42CD4661AA3267BBC00FE9BF033598486234087FA37B530580F8369E821C7350DCD48E067429E6E0C59D63801C404A0AF7EDFEFA7202616C9A9CA1488490A64A8A95740D1DC8ACADA061A40966A8CCAB4F17681FF463CABB1A00194337C42C2AB34FA350529E012107FFF00787A79C1CACA6A1020351315497B509E32DE08F86D404F95E920C06FD681E173E5D8F9E429AAEA1A68F3E10894F24C509A683EBE3614F9AE47F73D631AE0514F1F6C6C6CE175D19B2C3EA40B865220C81592B6CB8728BF1815350A1A80979A81921BA690C60F4680B253A146D558CBA4F0D23FD025FD1AFCEC1C54B576A1FC7E1B44772B905D5482942C016E262621F85A180DE5E6EE3194824644C527A138C684EC9AE306304447C16EC426F1212300B2879D34087DDCD6854279034465D5C82F97E376651DA4557528A8A8C1B5381EF2234D50106D36BE08E4FAACC5158B3F213C3E11D5C4D4C6C686EC865E60FBF92330E41A8223B988BC998084AC5CA40AF3915F518B420211F6130F92301348C24D8701369214C41229C791E91833050480B36F06014882BCED31AA1E3C82B8B402890211097302420840002714EE5EDE707673A75370BB5A81F01B09C8E51843186A8AB3CEF3A0E5A485900A47ACBDAAAFAC07E7111063D78021C22C6723F256126ADA9F40D6D2816212F622593DAD6299823E97D26A44496D23EED636D151C90E34C259F6C7B4F9CD7A36DA9ED7A0F0513C0C02E9A274FA05E27D00D1C775119D904C66DF8ED324052EA4D8BC7C2E2128240CAED7D9F8CCF70B1C8BB286A0B004658AFB28272B61348998E5616D68386A20A6D61DF7FBCA50F7AC104D7DA5F028FB1AEA8E2A4A881534C4D8006B71CB7E11226E25E35EE343DC21B31314DF411A693BDF1BC1D077358069AA01B6C7E961CDB97538627504B2A687381E6A0BB5B36A88949F47436F09E44FC5344058DD11ECCD53C77C7F2618762AD9EF055016619AF347E0F29251DDD28E6A120559731BF815795817B819DBE37571B17227FD62D36443EC0E37869F30049AE734C1A9B481ECA91095DDD9A8E9C94378FDF73011535814C454CE5E09707204C09B376FA8DEDE5E8A9F9A4A2DD35F4E31184C8AFC0D4741E02AC42424412C2D85B4528E2A4533B6857C018308067E289E8BF365EBE157FD0DB80DFFC5E1EC8DD03E3F059E77F7A1A48B87E2CE9BB8DD198FEB75DFE13FF913B1384869CC02D39EC565D8B3A68E0078FDFA35D5DFDF4F95DE29A52C2CCCA969D3B4A91CEF95288F5A0F71CC2164C5DA811FEF0B5E0A172A0E2AD89AC880994405C78A67C3A17405D8B25DB8566F097F99097E6AB24384E20704D598E35C99216DFEE15562FCA32A541DD4E3583FAA4D79A7085FBE78C1E8EBEDA11A1BEA9992421173E7F615AA99EE7A656511EB5015B301B2B88D28E36E4371F4BFB1D26D1E0DF14F026151C8C4BE2216F6DF5685A55483481396B727E2601151E12458164EC1128E1AD4CE4C80A6A356ACC699899ABFDB862DED0A86542A52E1A727B3D2D3F86A6E1E6E53C3ED96D9265F585A9FE9B6F4A5C86739A4C16B708F00E5461A6299EB2C1AE2F314168E954FC089AAC9B096CF809D7C0E4E57FD19D6E5BA3875672EF44334A179560B335C74E2B4CE694F1E75217ACCE150FBBFF2651C32DBC5DAFCE92E35C3F59FABAB686A4D5199C09C4B510CBDEFBF9C7322DC7AB128C969495B96BBDEAB44BF6558E53E17AC332AD89337096E8FE683DDB114FE8F5623A0ED637835ACC1DA88A9D0769A065DA7F969B3CECD9EF60A5DEF6E4AC36D686E6646696B6B53EAEA6A949A9AAA72644C9F389DA1BC4F51AA2A1463D204754A6D3A39F92B91C185BDF35C7C4E2F16AF725DF842DD9105EBEA05E0BDDE8DF4570791F4B3257626CCC30CE7593065EB414357632B79865589D0D1012C0F1DA27466EA1073758AA542FC984CFA03F3ED75E2162F87415E44560F6A029126D12C9DCD33D77FE0F6413DCB810576E367287AE582BD394B30DBF52F381ABC0EE1F61F8AC9EFFE46C4C887E3E800DC182E6566664EAD5CB5925AB87001355347E79DAFDCE1AFDEC94E53A939AE243314C524D258E9B97AE1521FFDAA991E5371386713743DE66387B7A13CFCCCDF8521A716ED1A8256C23347AD8196072D94244F42C5C5C5511C520F57AE5EA5823957A8ABC11C2A2323836A6E6EA6060606E8880C032CF226131BDADD367036CEFF24784BE1727F036CE16CB9B9E04B328BC128B194B31F827807E0FF76241F3BA0715CDB0000000049454E44AE426082;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 15;
			
		SET @maxid = @maxid + 1;
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)  
		SELECT @maxid, 'absPassword.png', 5, 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D0000002574455874646174653A63726561746500323031302D30332D32305431353A30323A34322D30353A30307B94C7050000002574455874646174653A6D6F6469667900323030392D31322D30315431303A34303A34382D30353A3030FA7322070000001974455874536F6674776172650041646F626520496D616765526561647971C9653C000006FB494441545847C5975B4C5BC919C779E843D58776DB4A512BA58F511FBA8D54B50F7989B46A57AD76B5BBD906020A59D810126EC96203013B5C0218B0B9047335EBC4E67EF182C3251408494C020448082640EC0502C1D83836B7A5812504081CFFFBCD51F0B2AABB3D525565A4BF66E6F83B33BF33DF37DF8CBD0078BD4DBDD5C9D9877B0478EF23EF5F9264241BA993D448CA2625923E237D403A4C3AF0BFAE9E47803F7F72422191656E3475DC7655D437EED4DF68DB5695566CE4AAD42F2FA5295645D2C47F0686452E1FF33FBD40B64ED210A993A4266593CE923E23FD9174F887203D02FCE598EF7A684C3CCE886211102EC231FF2004844622541C8B847405D272945069B4A8A8AB85BEB9112DFF68D9ADBF5EFF5AADD5BCCA57156F24CAD2D62449C9AB411191DF04849E5F7EFF533F2749E109C4230019EF361AFAD07C77004D5D7DB87EBB17E58DED50553520BD408D84AC7C44C6A7E25C9484073C1E100C6F5288281A62693C125265C8C953A2B45C8B5B9D37A0D296E2AFDEFE4EC1007FF3F6E7D8E456E722A6ACCFF164CA82CEBE4134DCBC879AD65B28BBDE06554D230A2AEA91ABAD4596BA02E9455A48E54A882FCB11214946B0488253619188888B477159053EF00D100EF0A16F20C756C053E1380E3BBB1C5EBE7A8595B575581D0B303DB5A0676814FA0E0354D50D8893E7E1627A2EA26539381E18024D8D0E1F9F3C2D1CE063FF204E7FABC73DBF42A1804C26E3FBABABAB78F2E48947B8FD0FB75EEF202A251B2782C250DBD88C4F03CE0A07F87BE0394ED771F77B00AC93949404B3D98C9898184C4C4CF0BF5B2C160C0F0FF3ED9D9D1D5E7B459C9285932117D072B3133E4161C2014E9C09E76A5AEFFCDB0A3000A954CA3FD768347C9D959585E0E060BEDDD3D303ABD5EA7E2F8A000223A260E8B90FBF731784039C0C89E4CA9B6EBA07AAAAAA72B719005B019148C43F0B0D0DC5FCFC3CDF562A95F0F7F777DBB2180816C5E1A17118A7C2C5C2010222A2394D43AB7BA0A1A12177DB6834F2ED888808BE1E191971AF808F8F0FEC76BBDB960562D8C50498C72770FA8B58E1004122095752DBE41E687F63636383EF767777F3754848088A8A8AF8368B8B3D77B0BE8476C3854BA9B0CC5A71363A5E3840C8C504AEB0B29E1F74616919BD838FD03DF0108F4D66586C7378B1B6E611CECFCF0F838383EEDFE2B38B109D9C01A7D389B0B8CBC20122A4291C4B309C0BE81F1A86637E0176E73C4C139330F4F6A35ADF84026D39E88CC0BDFE87304D4EC1B9B884ADADADEF815D5696202E2D1B4B4B4BB89090261C409494C1657E59816D02D0B77D178CFB4777B95C585E59C1D3190B1E8D8CC170BF1F2D9D77F87AF64D1CA4156A909099C7034425670A07889165732CB56E71C0F5F64E7E5E83C10087C30136F10F956FD7D7D1FFC8087A158A9272245F29E20162D3AF0807902AF2B8E43C353669146D9D9E9FEFC89123080F0FC7ECEC2CBCBDBDF9E4333333C3D72C3DEF2F6C156C9BC0154D0DD2E8F06200F15905C20112738AB9849C62ACEFD220574BF9B159E2292E2E86CD66C3A14387A0D3E95057578783070FF2D9502E9723232383B7BDF3601053AF4087D55750146BB14800977355C20152F3D59C54918F179455F3CBBE4B427B5FC95CC1227B7C7C1CE7CF9FC7DCDC1CC462318E1E3DCA9BDCA3C07D4A0025B58DC8B95A01E7C2226405D7840364149772B1194AAC10C0AD0743E8333EA6C07A8E35F2AF90D2FB780C93942E5832CBD3D6C0EE70425E52261C20EB6A2517436974F935B0485B61707286408C68B9DB0B3D45FABD8743181D9F84C5EEC0F28B17D8DCDE7673B120ED1B31618200CA9B3A50486EB0127CF6B52AA7D79B42C6EE8B90C71B516E692D17959A8D9AB63BB86F9EC6F437EB70D01CCF699BDB365D98A4FEA87D0943CF6C18F87A0AF7C7CC181835616CFA19CCCF2C30CE3A304E0035ADB7E9E2A2C7EC9C1DCAB23AE100399A6A734ABE9ABF583010B61A2CA994E89AA1EF1AC05DD334C616D7F08CFC3C4D62013749116B5AD9C0D8F23ABEA6C999D891AEA67718404165837080CFC323A313B30A560A2B1B5C45D57A468FECAB55CC8F74F5D2802518B6C7F3CA74A86AEB42EB8351743F9D8371E15B7E62F34BC0B4EE82FE760F1F070CA0B05ABF481EF831F3C27F75C11BC3DFBCFB873FBDF789EF29E939B1A44B927E655EAE2ADD2AAC6A70A9E8A06277421545F99775CDB8567F03A5744FAC6CE9441D7DF58DBE61743C32A189AE756574996500B4AAAF69DC44D23B4200BCC8DF5E76928D64DDA47F305E5E3F22FD94F4EE6F7FF7FB30DFCF835B45F129D3A979251BF9653A4EFD550BB47B10ED5DF4F5BD60F74A5DBB01ED14BC3E0141DBF46E26E9278200F62256407D806C0EFFEC9D9F1F7FFFA3636567BE8819BF24CF5DA32CB85B58A57729691B7EE8EDB77DE057BFD6B0AF17EA0201F3FE4793BD953A4C16FEA4545218E91782B7E1FE25FA7FB7DFFABFE37F01857E7738EF90A0510000000049454E44AE426082;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 18;
			
		SET @maxid = @maxid + 1;
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)  
		SELECT @maxid, 'absSignOut.png', 5, 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D0000002574455874646174653A63726561746500323031302D30332D32305431353A30323A34322D30353A30307B94C7050000002574455874646174653A6D6F6469667900323030392D31322D30315431303A34313A30302D30353A3030A21409A40000001974455874536F6674776172650041646F626520496D616765526561647971C9653C0000088F494441545847B597075055571AC75F36D9D164779CCDAE2EEA0A885D101124447A02642D58308ACB48EFBD88944713ECA2862010C8EA4AF1295D05C40A5808386261415869D2A5C3838702529EFFFDCE4D609E6B32B31BD93BF31BDEB9E79CFBFDCE77BE73EFC003C09B64246B3FEF7F4572FEAFF9CD05979595E558B06001879494D43C3D3DBD5D464646A74D4D4D1F4B4B4B0F13EDD417475810F327C7FE9AA09273FE53608E8CB4B4D99245722D0B172E848A8A0A4802E6E6E6D8BE7D3B141515212323030A7E9F906312D32930935629909393C3D2A54BA1A9A909333333B8B8B8C0C4C4846BAF58B102EBD7AFE7A4685C3E097C3C5D021FD0CA66129E6B5594C5AAAAAAD0D6D686818101F7575959192C23FAFAFAE0F3F9B0B0B080929212CBC4ECE912E051701E65809719623D767ABF374E86FA23F91FB1C8BB7219D9A90268ACFB1C8E8E8EB0B2B2829A9A1A14141498C09FA65DA0319E8FF18A1B186F7B06F14027C4833D98E86D45F6D928181A1A62E3C68DDC562C5BB6ECFF23D0941080B1D22C4C343E81B8A701E2BE564CB455A1E1560A0CF57440A703F2F2F29C006D8BD4B46780098C160B305E790B13F52598687A82F1EA4288F2E3116AB50D1BF574A1B26635D4D6280A690B3EB5B4B5E7BD0F53C770B20698C0EBFC188C162561EC5126C61E5FC2E8FD0B7875F1085E44BBE0B2D7166C505786FD56FD7212F8E87D82B3B9EF08343381ABE1787DF35B4E843172ED1B8812F6A03FD212077768B2238808A79DFBD87B60DA059AA80847B20F62E4CA618CE41EE578951688FE185B08BFD98D74CF6D70DEFAC590DD669DC5D32620F92DA83BEDD74FDF039238F0235961183CEB0EE1C9DDE88FB244C52957B49E0B0C99FC6658DA3A501D38D813F1C46309589BDD67FD94A99FE7AD57297B68799477E660662805260962282D0003D1D6DCEAEB22EDD1712108B87664031B6B65E7E8E6EEB5B72CFB4A6E5B754DCD2BBA30096BB3FBAC9F8D23D8F877F851407339077BE8DD63CEA6CFE2BC3194118457C97E10C5D973C15B22AC71EF9813DE5C3D82BB879CBC6DEC1C82CF27A736F5F5F54DB0A0EDED1D28FD6719F20B6EE3496929DADADA3819D69F9C9AD66C6DEFC427A4099E24EF085CDF6F375BE0B133ABF2A4DD7857940D7A23CC511B6E8182437618BAB49FAD1EF16E16B8919B3BC202343434E2E2E52C9C3D158982A4783C14C4E37A6C14BE3F7E0CA9E9E9A04C702279F9059D360ECE7C8227C99440EBA9A35C0672D5E5671ACBCEE92C3AEE349111608E343F5394C778D2761CE082E7C504222BF53CF7D0DB77EF22E2C409D464A54378231BBD99027427C6A2222E0A1D4703F1689F2F8E1D3C88FCDB77B8F11752D29A6C1D5DDC08DE246F095CD3505847A4D9CCFDC39B01019F0B28C970CE219CD81F00615F1FCA9F5620323C1CAD2989E84A8A437BCC71B41D0BC68B600F343F2A41A18B359A4D36A0CAD800C78383B9EDE9A5EDF0F2F52FB373769D45F0185302B4F219D73514EEB47FF5194E69A9206ABB365ACEF8BE2550FCF710E45ECC8048244292E002CABFA5975378085A42BCD0B4C7168DF6BBD068BA093D5D9D68AFADC5CDBF6AA1F9F3C578B0490B0949026EDEE5EC9C17F62EEE56048F312590A32EAFFBE42B35716BA80F0AFC3D106BBC05DF292DC67DB5A5A85BB704B56A8B7060A7112ACBCBF1BCBE1EB194DA066F3BD43B9AE0F9EE4D28F2714547F533B457FD0BC29E1E08BBBBD1FBF03E1E7839A27DAD0CA27D7D5057F71C55D5D5438EAE1E0282C79812B8BC6E6564B3D916502DA0E970102AFDDD51B84517B59A2BD040C1EB99808D253A3B3A70E75E21D27DDC51BD5903555F2AA2C8D71D038DF510F5F56270600083B4D2C1FE7E887ABA214C3F8704031D645BFE0D05540B9D5D5D622777CF6282C79812B8A4A5F4ACCDCB16ADA17BD1B5CF1B3D7C5774BB99A3CBD2089D5FEBA365A3068EB838A3ABAB0B57AFDF40A69D29AA741450A3BD122FEBAA313C28C2AB863A0C0BFB30323CCC31DC2F4482BE0E7A95E6A360D766E4E45EE5E67BEEF52B75F6D8338390D802ADD5E2125B13543B9942E8698D0117530C5A7F8D212AA4D7DB7430BE410D271D48B0B999ABEA54676BD46AADC473CDE5181DA0D5E6A4E3A6AE2A6EE9AEC5D8CB418CD13D2EF8EA79E85794C2357363DCCCCB476B6BABD8D5CBBB98E031A63210BF76794BA2CE5A94EDDA8497B4EA212AA611AAE0D16DDA5CF009FD353863B91BE58F1EE1694505628202F15C63395AB6EB62B4B9810BDEA6BA101D2A32186BAA478A9121FA56CF45FF2A298814E6E07BBE1F37AFACFCE990DB1E9F6482C7E004E8FA8DDBA279A6BECB1688B2B49430B2E34BBC36D2C198A13AC6D7AB72C1C55F28226BB33E2E25C4A3A9A90989740A8A8DF4B92A7F49ABCFD35141A78A34BAD7FC0589FADA102ACEC5C0AA3F63507E0E9EE9AB229E4E019B979679F185BBB7AF15C1634C0A7C4012B30DA43EB5B59695CA8B565E327A575D1EF5BA4A18D753C21BDD5518D7924781CA121CF170474D55158A8A8B1171E8106A7415D1EB6A86022D652EF864CA071458F0D9E85095437458180A7F28424D4D8D981F1256E6E9E33F8BE031A6B680047E4B48CDFAE84375C5599F381ACEFD63ACE3A2793FF82D976E8B5C253B12BD6A21FC97CC87F77A3D24C6C6A296CEF995ABD770F268380A8DE9EC53A149A6FCE5CAD92835D441C4E1A3DC38363EE97C72B3972FDF83E04DF2D6B7806D05F17B4296F88CD840181366842561C5D8B6C338233523535445997850528294F40CC4848522DDCD01372D8C91E96287B87DC11024A752A6EE838DA31750E71EBF003EC19364720B581D485E4CE413622EB1945843AC2334090D7905C5ADA65636DF9D894FEC282B2B13575656E22115277D7090957305B7A8DA4B1E3E04BBCFFA13CF9D6FF6F60FE413D2044F925F129094F9901A337FCA0CCBCEEF888F89199446EFA0B003E52969E96DF70A0B87CAE92D39096BB3FBD45FB6971FE441F07E8EFF46605286658515EB5B974F40F0A7843D21201E4BC0DAEC3EEBE7FD12EFFDCFA54F6008EF7DF8375EB6D110A2F1D37E0000000049454E44AE426082;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 17;
			
		SET @maxid = @maxid + 1;
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)  
		SELECT @maxid, 'absRefresh.png', 5, 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D00000025744558746372656174652D6461746500323030392D31312D32335431313A35383A31362D30353A30300FCF81FB00000025744558746D6F646966792D6461746500323030392D30382D32305431333A30353A32362D30353A3030A8E870FC0000001974455874536F6674776172650041646F626520496D616765526561647971C9653C00000A5B494441545847C5577957D3D716CD07A9EF23BCE754EB50AD56A64008210909843043984320400405874A05845AA8220E944920C818E61922B304641E4444E6190414BBFAA4FBDD7B5DF499BAFA77B3D65DF90DC9DDFB0CFB9C73390038FFE462E0068381D3D9D9C9696B6BE334353571EAEBEB39D5D5D59CF2F2724E494909A7A8A88813161A72581514E41F1CAC4A5787041B42D5EA75B53A785F15A45C0F08F033F8FAFAA4FBF878FBFBFA7A1FCECCCCE43C79F284939393C3C9CBCBE314161672743A1DA7A2A282EDADD7EBFF349A5D747474FCAD17349AF08B04F4F1B56B57B73333D3D1DCD488919161CCCFCF636D6D0D737373181C1C4443433D1E3F7A88884BE1DB9E9EEE690A8587C9DF79B6AEAECE98404B4BCB17042222345F858787DC8A8F8FDBA9ACAC208073D8DBDB63EBFDFBF7D8DDDDC5CECE0E5BEFDEBD63F7744D4F4FA3A4A418D1D1513B1EEE2EB7BCBCDCBEFA2B11EA898367ECE27346C9C9C91C8D26F45848B0AA243B3B1BAF5FBF66801B9B1B686A6BC2AFF969487C90004D742894017ED05C0945426A3CD2F21FA35E5F8B95D565466A7C7C1C0F1FA6C2D5455EE2E1EEFC35DDF700B4B8B8D898406565E59F0FD46AD5B1B030757379998E59FBE1C30794D79523E2070D82638210951E85585D2C929A9291D09C80CB6557A0CC5242714701E73047845E53A1A4AA186FDFBEC5D6D616B4DA3C7879BAEBDD5CE4C70F08E4E7E71B13282D2D650F8283830E0506F8E9CA0838055E5C5A40D2E33B50FDA044427102327A32F1A0EB211C726470C8258B7EE7C9E05EE081B0CA302875417088B787A3CA1E71776F616AEA153636369097970327B9BDCEC559F62F8A939B9B6B4C8066297DE1EFEF1B979595C1623ABF308F1F936FE2F2DD48687BF29165C8C243029ED2711F822C5BCC6CCD606A630A2D53CD486C4D807BB13B5C0B5D10591309A75FE5B00B13213AEE325EBE9CC0CACA0A52EE25C3D9C93E9EE26465651913D06AB504DCC7E4E68DEBBB9393938CC09D4789884E8D82B6371FE93D190CFC5E670A92DA92609B2D20FB80E5060DD3FAEE1ADA67DA115CA582245F0C5565109C9E3841A8B6C1CD9F6F607979192FFA7A71E9927AD7596E6F9A9191614C806AD55BE1F96B797919032FABD141131786024301320D9978D4FD18F7BBEE23A93D89596B9B23C0F2EE3236F736B1F7610FABABABD8DCDCC4EACE2AE2DB622129104155AD84E49118123F21720AB2B1B0B080DC9C6C383A4AD3D3D2D28C09F8FA781DBE1CA9D9999999C1DAFA1AC26FA8F1A0F601B27BB311511D89F0AA708456A91152158CF8560240AC143F15C28E00FDDC95888DBD0DCCCECD32C5AC6C2D23B62306B2520902ABFD611D63056FB5072626C6D1DFDF0FA5D26F472E971E3592A18FC22320EDF103667D5D730D226E8743DBAF455C4B1CC43962A474A5E06ED72F106985B8D4A0C10FFAEBB8AA8F42408D2FA42542D8EB84587EB784D1F151023481379BD3F0ADF5644B5A600781370FB9C40BAF5EBD4262421CE43289D28880B7B7477A7D5D2D2B2471F76E21E1690252BB52619F678FC8DA08FCFEF177FCF6DFDF087838A4452244B66870A93914214D415036FAC2AE8C0F55931F3676D7313030C0DCAD9DC8815BAD0C018D0AF0AE58101987B1DA9097FB0432A928D38880BF9FB7A1D76060DAF50B5320B5811490025778977AB3385349B2647BBF06D772396C8A2C21D659435066097E99196CCACD6153610EDD741126A72631343484FED53EB83448E1DE2083E51D3338FB3860747414B5B535B097087B8D080406F8AC8F8F8F31CD3AB93B4059100879BE232637261930AD6CB4A85082A3EB2310E808B08E001350DB6A73086B2DE0D06883D8FEEB585C5940476707E6366611D1AD82A88E0BEB2766104A79181E1E46575727247636EBC604FC7DF76902D1E62273102340178067D37A7CFCF8910153192D2D2DB11CF9E38F3F503B530941D5276051BD05EA17AAB0BFBFCFD4403A1D1A1B1BF1ECD9334658BFD448485AC056C065E179F1E205C422EB7D23027EBE5EEB3479D6D7D7217512C223D70D56595C96702BDB2B989E9D669BDD1FF805922A3EB35A546701712317764D5CC8F4D698D81E0569EBA42B3680B45C747777E3E5F6189CDB05B0ABE04260C7455F5F1F2326165A1B7B40E1E56E78FEBC9B69D935C011A2645B98647C0F59B114E36B6378F3E60D16161750F6BA18B635049C584D81ED9AB910357CBA5674CAB0B4B388D6D656E685155227FC7B9D21EBB282981863EF29424F4F0F74BA528884D606230FB8BB38A65790FA4FAD8C8E8F84E5253398665D8479AE094A5F16B1DEDF3FD08FB50FAB9FC009B0F49925242D8400B9A7DE10D65820B44D89D5CD55AC6DAEE1EAB01A8ECFAD20EFE141F0A33982A2FC98571E910E2914F08C55E028B30B484AFA89255B495521789EE630CBBE08AED68468D90BF36FE748F274312FDC1D4D84446F09FB364240CF651E115492457282124919FF090FA79218B093E1D3E2FB982125ED1746203A2A8210B032AE03523BC1113F6F8F1DDABD169717E11A48A413630A6EBE0978C5A6487A9180E9B96922A15A4CBD9E42EE743A543D9ECC13149812A0E0D42BB24E2B38F5F2E0DC6F0D97416B38645AC28194E3667D336A6A6AE0EA2CDBB115581E330A01ED4E7C9E797A4E4E16B6B7B7915D98012B0F3358A699C2BAD40CB655167834968237F3D32CC3A9CE5FADBDC4D8D6304636875818A4AD9670ECFE04EE3AC487FB281F2E24396D7CCD71FBFE2D9099138989F1E0599964D059D18800ED4E3C9EB9A9AF8FE7EED0D020B6DE6E41131B022B9529F8F966CC4A713D1731835718E8ECEC2C0B09CD78FAB16F27E0CF79702156BB8DF0E1316E03F73E3E44D116F0B9EEC6EA42696909DC5C64BB3C4B13D32F06123AC55246E6A617E2E2626398F6C727C7E01FA580750821A0250488E4A8959E3D12DC99BC89E6D51ABCDA9DC0DA6F2B70EEA3565BC3638C0FCF973670EBE0C3EE3A176E910EA8AAAB00197A1119110633D3736C1EF86224FB3F81F3872E9C3FA37B907A8FD58491B121A86F0582AF2045E7AE05EC698C89ACE424B128A8F38B4F71A6EEF69CB081170197E7F120545AC0FB860B19E574CCF53FC6DCC0C5F3A7756626E7D8444447F42F72E0E0C185EF4E1D3733F94E9F7237098B8B8B985B984352CE6DC8024510049AC3EE272E1C0A49AC5B79701BE4C3B58B90A8B4862C9507712879A712E06A9A066D5DADCCF29B37AFC1C4E49CFEFBF3A74F1C607C3E833226F41071F0924EAFDF9E397EFCCCC963A557A322D1DEDE86E937D3E81BE945D2D3DB08BEED0745840B9C1552482502387949E015298732C10BB179D751D356C9625E525C84307510BE3D7D4C77F6EC89139F4FC5444DC61EF89CC0019153DF1C3974F2C491589954B89B4C6A44072132323A8281E101740F76A265B0113503E5A8EBAB464B4FD3278B09303D4350979386B37BE6D4D1D8B3678E1F3AD8F3E09B28C998C0E753EA5F7F7CF2C47F4C4E7EFDEF34A18DE54E84464D2AD97DD0919DD6745A58A82C0B9E6A91F4732242540110F0B93BA74F1E493B73F288E95FF73AB827FFFD7228A524A827C861844DAD078BDED3E7248647898603256241A6AB9383C15E62BB4E6AC7BE99C9D9F50BE74E1988B599A7BE391C78FA9B2347E959929E290FCE835555551CEA76225B762EFCFC28F88F9E8CA947FE7102FF03A6FDE39C5D7573DD0000000049454E44AE426082;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 38;

		SET @maxid = @maxid + 1;
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)  
		SELECT @maxid, 'absCancel.png', 5, 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D00000025744558746372656174652D6461746500323030392D31312D32335431313A35383A31362D30353A30300FCF81FB00000025744558746D6F646966792D6461746500323030392D30382D32305431333A30353A30362D30353A3030EACD77810000001974455874536F6674776172650041646F626520496D616765526561647971C9653C00000952494441545847C55769545447167E2C82820A08C8A6C8D2A088AC2D5BA348088D0D8820821B1A511070016569DC30A8A080AD202891A889E8199D68D49818C7649C1333398E285BB76C822C4A2B2A6A0CD1B83BF9E6568384C52D333FE69DF39D57AFEA2EDFBD55B7AA1E0780FB7FE28DCE6B8A8A06FFAFC4AAB66FE7AA737339596EEE60E9B66D9C4C22E1A45BB628F0CAF66B09C876E66FBF549087CA820294482479FF2D91F29C1C49E5D61C54E6E4A03473C32715D99BB8F2AC2CAE6CF366CED4DC5CE1BB1F817249F60EA624DBBCF965B5241B15D9D9389791F1C59F2551B6614371C5C674C8D6AF476D7A3ACAD6AFC30FE2A4E2D2B434AE343DEDF5047E4C4DDD5BB66E1DAA3FFE18F70E1CC0AD63C7D098B509952B57FE5E929474FA7D4994A6A49C281727FD5EBB72256E1F3E8C3B84EA35AB70316939BE8D893E5A1A1FCF99595AF6CE40495C5C9E342606B50909682F2EC68D3D7B70A3A8086DD4BE9A9989CAF9F35F9C8B8C3CF92E12E5919147CAE7453C6D4E49C1CDFDFB1536AEEFDA85B67DFB70393616B2F070FC3334B450575FBF378133B367CF2D09103DAA0B0941DBEEDD901716429E9F8F56829C19D8B4093541414F2A84C22F2FFBFAAAF625522FF4512D15F97D5E1710F0AB5C2CC6F54F3F857CFBF64E908D360AE85A5212649E1ECFF3F9FC048E9E5E6B2066F162AB8290901B07A8BF5428847CDB365C235CA505A400B5DBB66C41EBD429CF4A5CF887DA2C2CBAABE416CF7C7089A35D61E364BF9FAFAF5D8B56727A8DD60E03D3BDBE6307EA29FAE31A1A28183BB6C38CC713F423A0ACA6A6B12024A4E8337B7BEC2112159327434E0E5B28F2968D1B3B41EDD68C0CDC0E0F7B79DE9A77F829C7693F52E1B42B4C47E4C943A775B4AE5A85E6CD9BBBE59B49E79A44822B717138A2A989A261C3103371E2B72AEAEAA64440A95706A843594D5373D4EC80803DC5AE2E285655414540005AD2D7A3891666635A5A375AE8BB79CEEC17673DDC8E9C13088A1AE6463CB84A0BAE91C9AD598346CA029367A4EB2223719422DF636C8C181F9FEF06EBE8B0E807F7CB80820DB1E2D4078E0C0B08D859ECE2FCF2A0A63AA4D3A7A389226BA0796DA085D5909CAC7837115A5353D14268A6B97D357685C9B071D2B91C15852FF5F4B0DBC00091DEDE5F0FD0D272231F9A0A3F7DD700EB78F5A80C1A6412EC27CCFDABEF041C35188A8A9933C8A11897972CC1E5A54B3BB16C19EAE3E31560ED6ED058C3F215A861911B19E0339E05667B791DD5D4D171ED8A5CE1FCAD0414A3AAAAFAFE4261D6215FF71727CD74218B5E882BC929A8591089DA458B501B1DDD09D67E058AB83E6E316A69CEBFE18DC23E4B53844DF03CA8ACA1E1481607F50CF2DD0498848AAA8EA7A767CA6E91D7DDEF6D8D50151F8BC6D495A88E8840F5BC79BD31772EEA8850F5E2589CB6B3C05E7BAB47BE2EE3F3953534EDC8D2C0BECEDF8F40A7D6C0712E2E31C553BCDAA5F3C3D0284EC5A5B0305C9A31A3376883A95B1885EA45F371C8CDE6B1B72B7F0BE95ABEC9793F029A9A9A9CA99111676460C019B3B7A121A7ABABCB0DA36F22E0F4B7ACACD6A645D1A8A4CA900607433A756A2758BB0B9553A6A07EE62C5C58BDAA43E0E616A7346000C7EC6A0D19C2E9E9E870BADADA9C0161C8A0CED9E855864E4E4E9C2445CC89E3E2B8D4C4442E39218113050773534343EDCEE5E4C85A3E9A8F4A3F3FC802033BC188D08655F9E1877FF4110129F537B00C89C5F717CC9C3977B48303E73B7122171112C2CD1289B83882608C4D7F020E249899B09C5B1E15C5AD58B64CC12E363A5A50969878A3392808526F6F5CA2CDE99248842A7F7FC8264D427D0CAD896431646CCCD757814A810017783CC8468D427D44C493D8A0A064662B3C30900B130AB928829BF5E83713584104962C5DCAA52725F9544545DDAB3131C145DAC14A478C40A9A929CACCCD71515F1FF5D3C3D171FE2C3ACEFD1D57139371415B07A546468AB10BBABA28D1D28294F6802B344DE9B36649442291DA5B09F43C5C32929226578587DFBD4C5148C9B1CCCC0C52D62602E564B869E61CDC3F751C0DB35C513DC104770E16419E244685CE3045E43222C95049F235F4DD4C9993444414FAFBFA0EECBE09F55D03AF0632626343690E9F345959A19A5037762C6A6D6C504BEF6A32268F8AC1FD6F8EA22ED80E3F390EC60F5E66A8F2D0C59D7DF9B8B93A0DB51696A81D3346A15347A8193D1A57C84E1B91C80A0CFC7CB69FDFA0AE5DB7FF8D68577CFCDCC6A0A0C7723B3B343A3AA2C9D9B9134E4E68A443EA66620AEEFD652F6A845638EB34ECC51C0FE72F84DEDE852726593FAC72D7C3EDC22CDCD998892607D2259D57FA8DD46E21FDDB3E3ED83D65CA117B9EB561BF2AE0F3F95E4D0BA31EDDF310E09AC003724F4FC8274CA0B700D76961DDDF9089BBBB0B207537C18F7CFD7F87B839EEE594551DC8D018BE93B3F82B9F71BF55D80FC1CDAC54FCB22D0F373EF0819CF41476BA70D3DD1D0F44FE98CFE7672A6968F4CEC0025BEBF5A58271B83F2D04ED7E22B45389DD26B4FB07E0418E04B7B2D6E2BCBD1ECE381B3EF71BEFF009A7A4E244CED92EC7A231B71D67977C6492DDC37FD96AA13539120FB6E5E26ED054B45365305BED425FDCF5F3470DDF069976234FAB6BEB28F7DA0756F32C8AD8E171D1D11C0F17C6A0636A303A68537992B7032D090B7186A787EF9C8D9F7B3A8C95D0C9CD1C9B11D409EC701946B01E61651D79C8D3B6E38C850EEAE74CC6E3DC3CFC3AEF23FC4265FC207406CAF8768A03AAC0C6B082535153EF26C00CA9AAA93B66F26DA54CA0D465349ED1827A9EBF1355A17EF87A843E4E391A3F771D37369B5352E6776DB16A3D2B87FAB4D9749858F0620E0AEC1E329D0A1F473CDD9283174B5740EA315EE1FCA4ABE973476BCBE59C92D2809E04944979B8FA106DBFADAEE31A9860A5970B2E4EF0C031E3E138E564FCCC638C4516053B9E45FEA68B298D69307226234DA3F7BBDBFCCC747F72B0A2CDC9BDD3B9B3F113576BF38D64C79DE434FBDE88582A0DD48768056D75B56D3C666C80AF4C0C71C2CDEC217FB4650629B9D0F84882CADB6EC65DD33246D7D8247297A7ED8DE364E338D93AE16CFC9B03CF6203D96119346253D7EFC7A46B3E0D070C1D1AB4DCCBF3FB7C77972ADEC8918924CB6E3206EFBA92F7D86458464DD5876A4D5B1728FAC7B1E0C9D71CCD4DD2C80EBB940CA76CB02AEC5D058A9E3F1E62A8E4C9A9A878519742A9CFF8FB7CAA2832A6A4E4C5A9A90BBBECB08A79CF1B516779B1634BEB7DBCBD4586916755A3D757E6B5FF863DD2C898BE75BEFFC4942856FCEBF01FC328F3CB1F769BD60000000049454E44AE426082;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 42;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 37;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 35;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 27;
					
		SET @maxid = @maxid + 1;
		INSERT INTO dbo.ASRSysPictures (PictureID, Name, PictureType, Picture)  
		SELECT @maxid, 'absOK.png', 5, 0x89504E470D0A1A0A0000000D4948445200000020000000200806000000737A7AF4000000017352474200AECE1CE90000000467414D410000B18F0BFC610500000006624B4744000000000000F943BB7F000000097670416700000020000000200087FA9C9D00000025744558746372656174652D6461746500323030392D31312D32335431313A35383A31362D30353A30300FCF81FB00000025744558746D6F646966792D6461746500323030392D30382D32305431333A30353A30342D30353A30307D5266A80000001974455874536F6674776172650041646F626520496D616765526561647971C9653C0000071F494441545847ED97695054571A864F2FD80DD242375B373D802DD04DD37D313324B1AC71CC68CA184140369545072D17508A8CA4E282B2C8D2E02846258120661242AC582975202AA6145121205182820BE2103480037121B88EF981EF9C73116C1613A72A8E7FE6C7577DEBF6ADFB3EDFF39D7B6E3701405E6611B24EF472EBFF002FD84024BD7F352DF639FAA85F2040CAB84C11DE2A1340962102CDC91C35CBFCE46FF534D07BA6D8648910729420EE9403D21BDE846A8B0583D836903198FB5B0398872FA91987031D9BD0F1E002F65E4B83365FC220F269B93E1700BD48F4DF58A1D76F1CE87C69AD0D0E756C41FBFD26B4DDAB47D7C31654757F8A3FEEB6198018C3430C90884CDC900D899E8F1724F1B38B7F1E9081F0B00A82A5B5B628EFDC8A6BF7CFA2F56E1DAD6F71F55E03BAFFDD8AEF6E7F05DFBFBB832459240E05C8340E02D02FD68E4D1761763981EE23E100B1F7B36CB0705B93102C7C59AD9CEF9C855FB95343AB960768BB77863F577CFE7388B27D40B20CB1A302B019B2551B7E5482C53563115545F06A89004F6C040E87300F8F3D658FAF3BB70D86B7DCF98687F8E7DD533C4051E34E88B327426032DE2526EEB5A1001B0D6C1C99AC93F00A0BA49D9B8CDC8B8188AF53F310D3F60920DEC08F24CE6C256F9467F777CE56FBD7D7B7F3F3BEDC5BD55F77AAF962160ACF15F09DD3F0FB34DC9356FFF8CD9E82D7ED7284985F6981F7EAF5C839FF163E688E40614B0CDE3DA3E521FC0F09C046C32C3158161E7E4C00D6399B39537DA9F7B8599D400B052838BB9D76EE03A189FB5990CDF9B2064603F864CA1E0B247CEB8AF50D7F4046E354DE40D1952528F9FE1DDEC8826A01422B0460A00A1A3EAF52400D2971988637F79E4053CF91C13AFFD3115CF8E918767CB7196336F9409CC3FD2CCCE6A60CDA1B6E40B2C9705E98628BA0034AACAEF746CAD949D4C20CEC689E8B4F5AE3F0E5B5246CBEF83616D74A115D2542E44911E24F2BF1558709E77ACA517FEB1FB44AD170FB005F8D3D87F17EBD09D2CD3E90E4707D6293D1CF7CFD8C306095EEFDD87A2BF750986C8D9083CED4C22BC868FA13B65CF243FE954814B7ADC0FE8E147A1C81B8D3B688AF77C2DEF6F5A8BBB907DFDCF80C35373EC7297A7CE6D63EFAA89521FBF43AD8BECF616C8EA16F4C9661DE88C53BDC8055A247A64DA61EF21DC6070C22AC5C8DD4C6D79073711AB6B70461E7F70B50F20385E85C8FCF7E88C517EDAB50F9E387A8E8CE436577014EFCB80BC7BA0AE8CE9783B88A50D8E5196193E3DD679DA11F11CEAF81AC618BD02A610291C56BF21CB37450E51B7B851B2C11546E83354D1A6434FB22B7753A0AAF85A2E4FA12ECED5A85FD5DABF1E5F5BF6277FB0A7CDCB610792DC1305D988A99FB75702BE4A0DCAC7F6C9BA65BF9CCBD2373188065EC78225BE64614CBDD3ED5E4EAA0F9D8D8234A9622A4528AC4CB32AC6D5560C35525D2DBC723ABDD13A6763D4C578DC86C3520ED32AD4B3E083830015EC51CC6E7EA1F2B933DD698878B53F54498662002F6B8A73F29F3C7D072912BB1A6258F5113A745EA529F7C2FE88A8DB7A569522C6DB04556B70B726F6991D7F30A3EEA998CA29E37B0F3E67414744D477EE70C441CF5C2AB7B3818B67941BD5A933ABC7371320548D113412A0D1F8030079046BB90B1512EC436D299DF925D16A88E4F2EF2C21BA5DC0D49AA04EBDADCB0EBE124EC7E3403FB1E85A2F45134CA1EC4A0ECFE32ACACF5C5B4520EAF7F40CD25B86D1D4DBB78AD8E08375080146F421804B3300460FEEF8855B833B109530DBE1326443855CFDCADC78C83DCBF2C374AB0B56B120EF645E3685F3C4EF6ADC5E9C7D9486E9A0AFFC31CFE5CA48376B9BAF0593317BF4701D679110135415268381DC75080306762354749C6CD711A04700D76106BA31C1BC30E1A1054C17531885DDDA1A843262EA01059CD33115C61847FB10EBAC5AA2F7EE9D52D4EF424426A4190440192A985FEADFFE9562C0D762696814E4416F01480DD501364676F8872B8BAF0880161955C8765DA18ECEB4E44CE653F449C34227C8F0EFA2887B25FFBDD205E4501A80501B540D6330B14620840908A58D27099BFE3883F2AEEB3152E13A31D3A975773987BDC78C5E16F3244561910BD5F0B5DB87DE5AF85B3EFC5EF7810E1BB5A2258430198851100014A22F57324325A8A59F6C4E96D3BA29E6547DC662988BB9F8278CE966B7D631C6FAEACE3B0A26162DFA2437AE843ED6A3C42EC897BA83DD1843912D7F94E443D5F451CA35444BE504D64312E44B2743C11C76A88388102249A01B031981B90D0EE7980591480863B32005A0C40E327279EFE72A29BADF8FD9404556F6489278C73EDCF7904D8893DE6D811770AA1097520AEE114609EB21F60819A5833802514603905887F02B05A4747402D0C3720097C62808E8019E001FC2800ED7EC25300A20F5484E802E4A5DA008593672035134C0D30803006E0480D288943643F80EC2F4F01440C601535C0009228C07003E6BF8EFFA7C7CFB3805EE4352FF5AF396BEC3FBEF7000147846E9A0000000049454E44AE426082;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 43;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 36;
		UPDATE tbsys_mobileformelements SET PictureID = @maxid WHERE ID = 28;
	END

	IF EXISTS(SELECT * FROM sys.columns WHERE Name = N'Width' and Object_ID = Object_ID(N'tbsys_mobileformelements')) 
	BEGIN
		ALTER TABLE dbo.tbsys_mobileformelements
		DROP COLUMN Width, Height, BackStyle, BackColor, HorizontalOffset, VerticalOffset, HorizontalOffsetBehaviour, VerticalOffsetBehaviour, PasswordType;
	END
	
	UPDATE [dbo].[tbsys_mobileformelements] SET Caption = 'Enter your registration details and an activation email will be sent to you.' WHERE ID = 44;
	UPDATE [dbo].[tbsys_mobileformelements] SET Caption = 'Forgot Username' WHERE ID = 13;
	UPDATE [dbo].[tbsys_mobileformelements] SET Caption = 'Enter your email address and an email will be sent to you confirming your username.' WHERE ID = 39;
		
	UPDATE [dbo].[tbsys_mobileformelements] SET FontBold = 0 WHERE Type = 3 AND FontBold  = 1;
	
	IF NOT EXISTS(SELECT * FROM [dbo].[tbsys_mobileformelements] WHERE ID = 50)
		EXEC sp_executesql N'INSERT [dbo].[tbsys_mobileformelements] ([ID], [Form], [Type], [Name], [Caption], [FontName], [FontSize], [FontBold], [FontItalic], [ForeColor], [PictureID]) 
		VALUES (50, 2, 2, N''lblNothingTodo'', N''You have nothing in your ''''action'''' list.'', N''Verdana'', 8.25, 0, 0, 0, NULL);';
	
	UPDATE [dbo].[tbsys_mobileformelements] SET Caption = 'You have no items in your list.' WHERE ID = 46;
	UPDATE [dbo].[tbsys_mobileformelements] SET Caption = 'You have no items in your list.' WHERE ID = 50;
	UPDATE [dbo].[tbsys_mobileformelements] SET Caption = 'Select an item to start.' WHERE ID = 49;
	UPDATE [dbo].[tbsys_mobileformelements] SET Caption = 'Select an item to continue.' WHERE ID = 47;
	
	UPDATE [dbo].[tbsys_mobileformelements] SET Name = 'btnCancel', Caption = 'Cancel' WHERE ID = 27;
	UPDATE [dbo].[tbsys_mobileformelements] SET Name = 'btnSubmit', Caption = 'OK' WHERE ID = 28;
	
	
	IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tbsys_mobilegroupworkflows]') AND type in (N'U'))
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[tbsys_mobilegroupworkflows](
			[UserGroupID] [int] NOT NULL,
			[WorkflowID] [int] NOT NULL,
			[Pos] [int] NOT NULL,
		 CONSTRAINT [PK_tbsys_mobilegroupworkflows] PRIMARY KEY CLUSTERED ([UserGroupID] ASC, [WorkflowID] ASC));';
	END

	IF NOT EXISTS(SELECT * FROM sys.columns WHERE Name = N'TodoTitleForeColor' and Object_ID = Object_ID(N'tbsys_mobileformlayout')) 
	BEGIN
		ALTER TABLE dbo.tbsys_mobileformlayout ADD
		TodoTitleForeColor int NULL,
		TodoDescForeColor int NULL,
		HomeItemForeColor int NULL
	END
	
	IF NOT EXISTS(SELECT * FROM sys.columns WHERE Name = N'TodoTitleFontUnderline' and Object_ID = Object_ID(N'tbsys_mobileformlayout')) 
	BEGIN
		ALTER TABLE dbo.tbsys_mobileformlayout ADD
			TodoTitleFontUnderline bit NULL,
			TodoTitleFontStrikeout bit NULL,
			TodoDescFontUnderline bit NULL,
			TodoDescFontStrikeout bit NULL,
			HomeItemFontUnderline bit NULL,
			HomeItemFontStrikeout bit NULL
			
		EXEC sp_executesql N'UPDATE dbo.tbsys_mobileformlayout SET 
			TodoTitleFontUnderline = 1,
			TodoTitleFontStrikeout = 0,
			TodoDescFontUnderline = 0,
			TodoDescFontStrikeout = 0,
			HomeItemFontUnderline = 1,
			HomeItemFontStrikeout = 0';
	END
	
	IF NOT EXISTS(SELECT * FROM sys.columns WHERE Name = N'FontUnderline' and Object_ID = Object_ID(N'tbsys_mobileformelements')) 
	BEGIN
		ALTER TABLE dbo.tbsys_mobileformelements ADD
		FontUnderline bit NULL,
		FontStrikeout bit NULL
				
		EXEC sp_executesql N'UPDATE dbo.tbsys_mobileformelements SET 
			FontUnderline = 0,
			FontStrikeout = 0';
	END
	

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
			SELECT TOP 5 E.type,
				S.instanceID,
				E.ID,
				S.ID
			FROM ASRSysWorkflowInstanceSteps S
			INNER JOIN ASRSysWorkflowElements E ON S.elementID = E.ID
			WHERE S.status = 1
				AND E.type <> 5 -- 5 = StoredData elements handled in the service
			ORDER BY s.ActivationDateTime;
			
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
					EXEC [dbo].[spASRSubmitWorkflowStep] @iInstanceID, @iElementID, '''', @sForms OUTPUT, @fSaveForLater OUTPUT, 0;
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
			WHERE ((WIS.status = 2) OR (WIS.status = 7)) -- Pending user action/completion
				AND isnull(WE.timeoutFrequency,0) > 0
				AND CASE 
						WHEN WE.timeoutPeriod = 0 THEN 
							dateadd(minute, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 1 THEN 
							dateadd(hour, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 2 AND WE.timeoutExcludeWeekend = 1 THEN 
							dbo.udfASRAddWeekdays(WIS.activationDateTime, WE.timeoutFrequency)
						WHEN WE.timeoutPeriod = 2 THEN 
							dateadd(day, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 3 THEN 
							dateadd(week, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 4 THEN 
							dateadd(month, WE.timeoutFrequency, WIS.activationDateTime)
						WHEN WE.timeoutPeriod = 5 THEN 
							dateadd(year, WE.timeoutFrequency, WIS.activationDateTime)
						ELSE getDate()
					END <= getDate();	
		
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
	-- spASRInstantiateTriggeredWorkflows
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRInstantiateTriggeredWorkflows]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRInstantiateTriggeredWorkflows]
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
					parent2RecordID,
					pageno)
				VALUES (@iWorkflowID, 
					@iRecordID, 
					0, 
					''<Triggered>'',
					@iParent1TableID,
					@iParent1RecordID,
					@iParent2TableID,
					@iParent2RecordID,
					0)
								
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
				
				FETCH NEXT FROM triggeredWFCursor INTO @iQueueID, @iRecordID, @iWorkflowID, @iParent1TableID, @iParent1RecordID, @iParent2TableID, @iParent2RecordID
			END
			CLOSE triggeredWFCursor
			DEALLOCATE triggeredWFCursor
		END
		';

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

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRInstantiateWorkflow]
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
				@fSaveForLater			bit,
				@fResult		bit;
		
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
				[parent2RecordID],
				[pageno])
			VALUES (@piWorkflowID, 
				@iInitiatorID, 
				0, 
				@sActualLoginName,
				@iParent1TableID,
				@iParent1RecordID,
				@iParent2TableID,
				@iParent2RecordID,
				0);
						
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
							@iStoredDataRecordID	OUTPUT,
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

	EXECUTE sp_executeSQL @sSPCode;

	----------------------------------------------------------------------
	-- spASRMobileInstantiateWorkflow
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
						@fSaveForLater			bit,
						@fResult	bit;
				
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
						[parent2RecordID],
						pageno)
					VALUES (@piWorkflowID, 
						@iInitiatorID, 
						0, 
						@sActualLoginName,
						@iParent1TableID,
						@iParent1RecordID,
						@iParent2TableID,
						@iParent2RecordID,
						0);
								
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
									@iStoredDataRecordID	OUTPUT,
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
				
				END;';

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
				@pfSavedForLater	bit				OUTPUT,
				@piPageNo	integer
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
						
						/* Remember the page number too  */
						UPDATE ASRSysWorkflowInstances
						SET ASRSysWorkflowInstances.pageno = @piPageNo
						WHERE ASRSysWorkflowInstances.ID = @piInstanceID;
		
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
							ASRSysWorkflowInstances.status = 3,
							ASRSysWorkflowInstances.pageno = @piPageNo
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
			END';

	EXECUTE sp_executeSQL @sSPCode;

		

	----------------------------------------------------------------------
	-- spASRSysMobilePasswordOK Stored Procedure
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSysMobilePasswordOK]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].spASRSysMobilePasswordOK;

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRSysMobilePasswordOK]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;
	
	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRSysMobilePasswordOK]
		(
			@sCurrentUser VARCHAR(MAX)
		)
		AS
		BEGIN
			/* Update the current user''s record into ASRSysPassword table.. */
			DECLARE @iCount		integer;		
			/* Check that the current user has a record in the table. */
			SELECT @iCount = COUNT(userName)
			FROM ASRSysPasswords
			WHERE userName = @sCurrentUser;
			IF @iCount = 0
			BEGIN
				INSERT INTO ASRSysPasswords (userName, lastChanged, forceChange)
				VALUES (@sCurrentUser, GETDATE(), 0);
			END
			ELSE
			BEGIN
				UPDATE ASRSysPasswords 
				SET lastChanged = GETDATE(), 
					forceChange = 0
				WHERE userName = @sCurrentUser;
			END
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

		IF NOT @tablename = ''''
		BEGIN
			SET @ssql = ''SELECT @ID = MAX(ID) FROM '' + @tablename;
			EXECUTE sp_executesql @ssql, N''@ID int OUTPUT'', @ID = @piNewRecordID OUTPUT;
		END
		ELSE SET @piNewRecordID = 0	
		
END'

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
    DECLARE @iTimestamp integer,
			@sSQL		nvarchar(MAX);

	-- Get status of amended record
	EXEC dbo.sp_ASRRecordAmended @piResult OUTPUT,
	    @piTableID,
		@psRealSource,
		@piID,
		@iTimestamp;

	-- If Ok run the delete statement
    IF @piResult <> 3
    BEGIN
       SET @sSQL = ''DELETE '' +
            '' FROM '' + @psRealSource +
            '' WHERE id = '' + convert(varchar(MAX), @piID);
       EXECUTE sp_executesql @sSQL;
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

	-- Get status of amended record
	EXEC dbo.sp_ASRRecordAmended @piResult OUTPUT,
	    @piTableID,
		@psRealSource,
		@piID,
		@piTimestamp;

    -- Run the given SQL UPDATE string.   
    IF @piResult = 0
		EXECUTE sp_executeSQL @psUpdateString;

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
PRINT 'Step 10 - Fusion Services (may be superseded by Fusion Installer)'

	IF NOT EXISTS(SELECT * FROM sys.schemas where name = 'fusion')
		EXECUTE sp_executesql N'CREATE SCHEMA [fusion];';

	---- Enable the service broker
	--IF NOT EXISTS(SELECT is_broker_enabled FROM sys.databases WHERE is_broker_enabled  = 1 AND name = @DBName)
	--BEGIN
	--	SET @NVarCommand = 'ALTER DATABASE [' + @DBName + '] SET NEW_BROKER';
	--	EXEC sp_executeSQL @NVarCommand;
	--END

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


	-- Create fusion core tables
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

	IF NOT EXISTS(SELECT * FROM sys.sysobjects where name = 'MessageDefinition' AND xtype = 'U')
		EXECUTE sp_executesql N'CREATE TABLE fusion.[MessageDefinition] (
			[ID] smallint NOT NULL,
			[Name] varchar(255) NOT NULL,
			[Description] varchar(MAX) NOT NULL,
			[Version] tinyint NOT NULL,
			[AllowPublish] bit NOT NULL, 
			[AllowSubscribe] bit NOT NULL,			
			[TableID] integer NULL,
			[StopDeletion] bit NOT NULL,
			[BypassValidation] bit NOT NULL
			CONSTRAINT [PK_MessageCategory] PRIMARY KEY CLUSTERED ([id] ASC));';

	IF NOT EXISTS(SELECT * FROM sys.sysobjects where name = 'ValidationWarnings' AND xtype = 'U')
		EXECUTE sp_executesql N'CREATE TABLE fusion.[ValidationWarnings] (
			[ID] integer IDENTITY(1,1) NOT NULL,
			[TableID] smallint NOT NULL,
			[RecordID] integer NOT NULL,
			[MessageName] varchar(255) NOT NULL,
			[ValidationMessage] varchar(MAX),
			[CreatedDateTime] datetime NOT NULL
			CONSTRAINT [PK_ValidationWarnings] PRIMARY KEY CLUSTERED ([ID] ASC));';

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
	IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[fusion].[spGetMessageDefinitions]') AND type in (N'P', N'PC'))
		DROP PROCEDURE [fusion].[spGetMessageDefinitions];

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

	EXECUTE sp_executesql N'CREATE PROCEDURE fusion.[spGetMessageDefinitions]
	AS
	BEGIN
		SELECT [ID], [name], [description],
			[version], [allowpublish], [allowsubscribe], [bypassvalidation], [stopdeletion],
			ISNULL([tableid],0) AS [tableid]
			
		 FROM fusion.[MessageDefinition]
	END';

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
			FROM SERVICE FusionApplicationService 
			TO SERVICE ''FusionConnectorService''
			ON CONTRACT TriggerFusionContract
			WITH ENCRYPTION = OFF;
		
		DECLARE @msg varchar(max);

		SET @msg = (SELECT	@MessageType AS MessageType, 
							@LocalId as LocalId,
							CONVERT(varchar(50), GETUTCDATE(), 126)+''Z'' as TriggerDate 
						FOR XML PATH(''SendMessage''));	
		
		SEND ON CONVERSATION @DialogHandle
			MESSAGE TYPE TriggerFusionSend (@msg);
	 
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
/* Attempts to default some new configuration options.           */
/* ------------------------------------------------------------- */
PRINT 'Step 12 - Auto Configuration'

	SELECT @perstableid = ISNULL([parametervalue],0) FROM dbo.[ASRSysModuleSetup]
		WHERE ModuleKey = 'MODULE_PERSONNEL' AND ParameterKey = 'Param_TablePersonnel' AND [ParameterType] = 'PType_TableID';

	IF @perstableid > 0
	BEGIN

		EXECUTE dbo.spstat_setdefaultmodulesetting 'MODULE_MOBILE', 'Param_TablePersonnel', @perstableid, 'PType_TableID';

		SET @columnid = 0;
		EXECUTE dbo.spstat_scriptnewcolumn @columnid OUTPUT, @perstableid, 'Mobile_Workflow_Activated', -7, 'Checkbox if the workflow mobile is activated for this user', 1, 0, 1, '5371A8DF-39BC-4A8E-875D-0DADE806F0BA';
		--EXECUTE dbo.spstat_setdefaultmodulesetting 'MODULE_MOBILE', 'Param_MobileActivated', @columnid, 'PType_ColumnID';

	END

/* ------------------------------------------------------------- */
/* Create Mobile Licensing SP.           */
/* ------------------------------------------------------------- */
PRINT 'Step 13 - Additional Mobile Configuration'
	----------------------------------------------------------------------
	-- spASRSysGenMobileLicence
	----------------------------------------------------------------------

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRSysGenMobileLicence]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRSysGenMobileLicence];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRSysGenMobileLicence]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRSysGenMobileLicence](
		  @piLicenceQty bigint
		  ) 
		AS
		BEGIN
		  DECLARE 
		  @sGUID varchar(MAX),
		  @iCount integer;
		
			SELECT @iCount = COUNT(*) FROM ASRSysSystemSettings WHERE [Section] = ''licence'' AND [SettingKey] = ''mobile'';
			
			IF @iCount > 0 DELETE FROM ASRSysSystemSettings WHERE [Section] = ''licence'' AND [SettingKey] = ''mobile'';
		
			IF @piLicenceQty <= 0 RETURN;
		
			SET @sGUID = NEWID();
			
			SET @sGUID = @sGUID + ''-EA'' + CONVERT(VARCHAR(MAX), @piLicenceQty) + ''FF'';
			
			INSERT INTO ASRSysSystemSettings
				(Section, SettingKey, SettingValue)
				VALUES 
				(''licence'', ''mobile'', @sGUID);
		
		END;';

	EXECUTE sp_executeSQL @sSPCode;


/* ------------------------------------------------------------- */
/* Create Mobile Licensing SP.           */
/* ------------------------------------------------------------- */
PRINT 'Step 14 - Legacy Data Cleansing'

	-- Orphaned order objects
	EXECUTE sp_executeSQL N'DELETE FROM ASRSysOrderItems
		WHERE OrderID NOT IN (SELECT OrderID FROM ASRSysOrders);';






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
	EXEC spsys_setsystemsetting 'outlook service', 'minimum version', '5.0.0';
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
