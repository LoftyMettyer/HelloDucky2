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
	
DECLARE @ownerGUID uniqueidentifier
DECLARE @nextid integer
DECLARE @sSPCode nvarchar(max)

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
IF (@iSQLVersion < 10)
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

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersAppName]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersAppName];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetCurrentUsersCountOnServer]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetCurrentUsersCountOnServer];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRGetDomainPolicy]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRGetDomainPolicy];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRSubmitWorkflowStep]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spASRSubmitWorkflowStep];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_audittable]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_audittable];

	IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spstat_getaudittrail]') AND xtype = 'P')
		DROP PROCEDURE [dbo].[spstat_getaudittrail];

		

	EXECUTE sp_executeSQL N'CREATE PROCEDURE [dbo].[spstat_setdefaultsystemsetting](
			@section AS varchar(50),
			@settingkey AS varchar(50),
			@settingvalue AS nvarchar(MAX))
		AS
		BEGIN
			IF NOT EXISTS(SELECT [SettingValue] FROM [asrsyssystemsettings] WHERE [Section] = @section AND [SettingKey] = @settingkey)
				INSERT ASRSysSystemSettings([Section], [SettingKey], [SettingValue]) VALUES (@section, @settingkey, @settingvalue);	
		END';


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
				FROM dbo.ASRSysAuditTrail'';

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
				FROM dbo.ASRSysAuditPermissions'';

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
				FROM dbo.ASRSysAuditGroup'';

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
				FROM dbo.ASRSysAuditAccess'';

			IF LEN(@psOrder) > 0
				SET @sExecString = @sSQL + @psOrder;
			ELSE
				SET @sExecString = @sSQL;

		END

		-- Retreive selected data
		IF LEN(@sExecString) > 0 EXECUTE sp_executeSQL @sExecString;

	END'




/* ------------------------------------------------------------- */
PRINT 'Step 2 - Set default values'

	EXEC dbo.[spstat_setdefaultsystemsetting] 'integration', 'payroll', 'OpenPay';



/* ------------------------------------------------------------- */
PRINT 'Step 3 - Product Rebrand'




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
	EXEC spsys_setsystemsetting 'workflow service', 'minimum version', '4.2.0';
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
