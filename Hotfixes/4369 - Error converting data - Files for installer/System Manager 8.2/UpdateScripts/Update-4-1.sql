
/* --------------------------------------------------- */
/* Update the database from version 4.0 to version 4.1 */
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
IF (@sDBVersion <> '4.0') and (@sDBVersion <> '4.1')
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
PRINT 'Step 1 - Modifying Workflow procedures'

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
				@fIsDelegate		bit,
				@sTemp		varchar(MAX),
				@fCopyDelegateEmail		bit;
		
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
						''HR Pro Workflow'',
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
					exec [dbo].[spASREmailImmediate] ''HR Pro Workflow'';
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
								ELSE convert(varchar(MAX), @sResult)
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
	EXECUTE sp_executeSQL @sSPCode;







/* ------------------------------------------------------------- */
PRINT 'Step 2 - Version 1 Integration Modifications'


	-- Create document management map table
	IF OBJECT_ID('ASRSysDocumentManagementTypes', N'U') IS NULL	
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysDocumentManagementTypes]
                    ( [DocumentMapID]			integer			NOT NULL IDENTITY(1,1)
                    , [Name]					nvarchar(255)
                    , [Description]				nvarchar(MAX)
                    , [Access]					varchar(2)
                    , [Username]				varchar(50)
                    , [CategoryRecordID]		integer
                    , [TypeRecordID]			integer                    
                    , [TargetTableID]			integer
                    , [TargetKeyFieldColumnID]	integer
                    , [TargetColumnID]			integer
                    , [TargetCategoryColumnID]	integer
                    , [TargetTypeColumnID]		integer
                    , [TargetGUIDColumnID]		integer                    
                    , [Parent1TableID]			integer
                    , [Parent1KeyFieldColumnID]	integer
                    , [Parent2TableID]			integer
                    , [Parent2KeyFieldColumnID]	integer
                    , [ManualHeader]			bit
                    , [HeaderText]				nvarchar(MAX))
               ON [PRIMARY]'
	END	


	-- Add columns to ASRSysControls
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysControls', 'U') AND name = 'NavigateTo')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysControls
								ADD [NavigateTo] nvarchar(MAX), [NavigateIn] tinyint, [NavigateOnSave] bit';
	END	

	-- Insert the system permissions for Document Management
	IF NOT EXISTS(SELECT * FROM dbo.[ASRSysPermissionCategories] WHERE [categoryID] = 43)
	BEGIN
		INSERT dbo.[ASRSysPermissionCategories] ([CategoryID], [Description], [ListOrder], [CategoryKey], [Picture])
			VALUES (43, 'Document Types', 10, 'VERSION1',0x00);
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (154,43,'New', 10, 'NEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (155,43,'Edit', 20, 'EDIT');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (156,43,'View', 30, 'VIEW');
		INSERT dbo.[ASRSysPermissionItems] ([ItemID], [CategoryID], [Description], [ListOrder], [ItemKey])
			VALUES (157,43,'Delete', 40, 'DELETE');
	END

	-- Update the system permission image for Document Management
	IF EXISTS(SELECT * FROM dbo.[ASRSysPermissionCategories] WHERE [categoryID] = 43)
	BEGIN
		SELECT @ptrval = TEXTPTR([picture]) 
		FROM dbo.[ASRSysPermissionCategories]
		WHERE categoryID = 43;

		WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x424D360300000000000036000000280000001000000010000000010018000000000000030000C40E0000C40E00000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF404040404040404040404040404040404040404040404040404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000005F5F5F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF9F9F9FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF7F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4040404040404040404040402020204040404040404040402020207F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFFFFFFFFFFFFF7F7F7F5F5F5FB0B0B0C0C0C04040407F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFBFBFBFBFBFBF5F5F5FFFFFFF5F5F5FB0B0B04040407F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFFFFFFFFFFFFF7F7F7FFFFFFFFFFFFF5F5F5F4040407F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFBFBFBFBFBFBF9F9F9F7F7F7F7F7F7F7F7F7F2020207F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFFFFFFF4040407F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4040407F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFFFFFFF4040407F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFFFFFFF4040407F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF4040407F7F7F404040FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFBFBFBFBFBFBFBFBFBFBFBFBFBFBFBFFFFFFF4040407F7F7FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFA0A0A0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF404040909090FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF909090A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0A0606060FFFFFFFFFFFFFFFFFFFFFFFF00
	END


/* ------------------------------------------------------------- */
PRINT 'Step 3 - Mail Merge Modifications'


	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'OutputFormat')
    BEGIN

    	IF EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'OutputPrinterName')
			EXEC sp_executesql N'ALTER TABLE ASRSysMailMergeName DROP COLUMN [OutputPrinterName]';

    	IF EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'DocumentMapID')
			EXEC sp_executesql N'ALTER TABLE ASRSysMailMergeName DROP COLUMN [DocumentMapID]';

    	IF EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'ManualDocManHeader')
			EXEC sp_executesql N'ALTER TABLE ASRSysMailMergeName DROP COLUMN [ManualDocManHeader]';

		EXEC sp_executesql N'ALTER TABLE [ASRSysMailMergeName] ADD
				[OutputFormat]       [int] NULL,
				[OutputScreen]       [bit] NULL,
				[OutputPrinter]      [bit] NULL,
				[OutputPrinterName]  [varchar] (255) NULL,
				[OutputSave]         [bit] NULL,
				[OutputFilename]     [varchar] (255) NULL,
				[DocumentMapID]      integer NULL,
				[ManualDocManHeader] bit NULL';

		EXEC sp_executesql N'UPDATE [ASRSysMailMergeName] SET
				[OutputFormat]       = CASE WHEN Output=2   THEN 1 ELSE 0 END,
				[OutputScreen]       = CASE WHEN CloseDoc=0 THEN 1 ELSE 0 END,
				[OutputPrinter]      = CASE WHEN Output=1   THEN 1 ELSE 0 END,
				[OutputPrinterName]  = '''',
				[OutputSave]         = DocSave,
				[OutputFileName]     = DocFileName,
				[DocumentMapID]      = 0,
				[ManualDocManHeader] = 0';

		EXEC sp_executesql N'ALTER TABLE [ASRSysMailMergeName] DROP COLUMN
				[Output],
				[CloseDoc],
				[DocSave],
				[DocFileName]';

		EXECUTE spASRResizeColumn 'ASRSysMailMergeName','EmailSubject','MAX';
		EXECUTE spASRResizeColumn 'ASRSysMailMergeName','EmailAttachmentName','MAX';

	END

	EXEC sp_executesql N'UPDATE [ASRSysMailMergeName] SET
			      [OutputScreen] = 1
			WHERE [OutputFormat] = 0
			  AND [OutputPrinter] = 0 AND [OutputSave]= 0';


/* ------------------------------------------------------------- */
PRINT 'Step 3 - Office Output Formats'


	EXEC sp_executesql N'IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[ASRSysFileFormats]'') AND type in (N''U''))
		DROP TABLE [dbo].[ASRSysFileFormats]'

	EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysFileFormats]
		( [ID] [int] NULL
		, [Destination] [varchar](255) NULL
		, [Description] [varchar](255) NULL
		, [Extension] [varchar](255) NULL
		, [Office2003] [int] NULL
		, [Office2007] [int] NULL
		, [Default] [bit] NULL
		) ON [PRIMARY]'

	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(901,''Word'',''Word 97-2003 Document (*.doc)''        ,''doc'',     0, 0,1)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(902,''Word'',''Word Document (*.docx)''               ,''docx'', null,16,0)'
	--EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(903,''Word'',''XML document format (*.xml)''          ,''xml'',  null,12,0)'
	--EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(904,''Word'',''PDF format (*.pdf)''                   ,''pdf'',  null,17,0)'
	--EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(905,''Word'',''XPS format (*.xps)''                   ,''xps'',  null,18,0)'

	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(911,''WordTemplate'',''Word 97-2003 Document (*.doc)'',''doc'',     0, 0,0)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(912,''WordTemplate'',''Word 97-2003 Template (*.dot)'',''dot'',     0, 1,1)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(913,''WordTemplate'',''Word Document (*.docx)''       ,''docx'', null, 0,0)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(914,''WordTemplate'',''Word Template (*.dotx)''       ,''dotx'', null, 14,0)'

	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(921,''Excel'',''Excel 97-2003 Workbook (*.xls)'',''xls'', -4143,56,1)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(922,''Excel'',''Excel Workbook (*.xlsx)''       ,''xlsx'', null,51,0)'

	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(921,''ExcelTemplate'',''Excel 97-2003 Template (*.xlt)'',''xlt'', 17,17,1)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(922,''ExcelTemplate'',''Excel Template (*.xltx)''       ,''xltx'', null,17,0)'


/* ------------------------------------------------------------- */
PRINT 'Step 4 - Overlapping dates functionality'

	IF OBJECT_ID('ASRSysTableValidations', N'U') IS NULL	
	BEGIN
		EXEC sp_executesql N'CREATE TABLE [dbo].[ASRSysTableValidations](
			[ValidationID]				integer NOT NULL,
			[TableID]					integer NOT NULL,
			[Type]						tinyint NOT NULL,
			[EventStartDateColumnID]	integer,
			[EventStartSessionColumnID] integer,
			[EventEndDateColumnID]		integer,
			[EventEndSessionColumnID]	integer,
			[FilterID]					integer,
			[Severity]					tinyint,
			[Message]					nvarchar(MAX)
		 CONSTRAINT [PK_ASRSysTableValidations] PRIMARY KEY CLUSTERED 
		([ValidationID] ASC)) ON [PRIMARY]'
	END

	-- Add columns to ASRSysTableValidations
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysTableValidations', 'U') AND name = 'EventTypeColumnID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysTableValidations
								 ADD [EventTypeColumnID] integer';
	END

	-- Add columns to ASRSysTableValidations
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysTableValidations', 'U') AND name = 'ColumnID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysTableValidations
								 ADD [ColumnID] integer, [ValidationGUID] uniqueidentifier';
	END


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[udfASRDateOverlap]')
			AND OBJECTPROPERTY(id, N'IsScalarFunction') = 1)
		DROP FUNCTION [dbo].[udfASRDateOverlap];

	SET @sSPCode = 'CREATE FUNCTION dbo.[udfASRDateOverlap]()
		RETURNS bit
		AS
		BEGIN
			DECLARE @bDummy bit;
			RETURN @bDummy;
		END';
	EXECUTE sp_executeSQL @sSPCode;


	SET @sSPCode = 'ALTER FUNCTION dbo.[udfASRDateOverlap](
		@pdStartDate1		datetime,
		@psStartSession1	nvarchar(2),
		@pdEndDate1			datetime,
		@psEndSession1		nvarchar(2),
		@psType1			nvarchar(MAX),
		@pdStartDate2		datetime,
		@psStartSession2	nvarchar(2),
		@pdEndDate2			datetime,
		@psEndSession2		nvarchar(2),
		@psType2			nvarchar(MAX))
	RETURNS bit
	AS
	BEGIN
		
		DECLARE @bFound bit;
		SET @bFound = 0;
		
		-- 1st of data is the inserted, 2nd is physical database values.	
		SET @pdStartDate1 = DATEADD(D, 0, DATEDIFF(D, 0, @pdStartDate1));
		SET @pdEndDate1 = ISNULL(DATEADD(D, 0, DATEDIFF(D, 0, @pdEndDate1)), CONVERT(datetime,''9999-12-31''));
		SET @pdStartDate2 = DATEADD(D, 0, DATEDIFF(D, 0, @pdStartDate2));
		SET @pdEndDate2 = ISNULL(DATEADD(D, 0, DATEDIFF(D, 0, @pdEndDate2)), CONVERT(datetime,''9999-12-31''));

		-- Put the AM/PM stuff into the above dates.
		IF @psStartSession1 = ''PM'' SET @pdStartDate1 = DATEADD(hh, 12, @pdStartDate1);
		IF @psEndSession1 = ''PM'' SET @pdEndDate1 = DATEADD(hh, 23, @pdEndDate1);
		IF @psStartSession2 = ''PM'' SET @pdStartDate2 = DATEADD(hh, 12, @pdStartDate2);
		IF @psEndSession2 = ''PM'' SET @pdEndDate2 = DATEADD(hh, 23, @pdEndDate2);

		-- Check to see if this date overlaps.
		IF ((@pdStartDate1 BETWEEN @pdStartDate2 AND @pdEndDate2)
			OR (@pdEndDate1 BETWEEN @pdStartDate2 AND @pdEndDate2)
			OR (@pdStartDate1 < @pdStartDate2 AND @pdEndDate1 > @pdEndDate2))
			AND (@psType1 = @psType2 OR @psType1 IS NULL)
				SET @bFound = 1;

		RETURN @bFound;

	END'
	EXECUTE sp_executeSQL @sSPCode;





/* ------------------------------------------------------------- */
PRINT 'Step 5 - Intranet Dashboard Implementation'

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'SeparatorOrientation')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD SeparatorOrientation int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET SeparatorOrientation = 0'
	END
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'PictureID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD PictureID int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET PictureID = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_Type')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_Type int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_Type = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_ShowLegend')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_ShowLegend bit NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_ShowLegend = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_ShowGrid')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_ShowGrid bit NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_ShowGrid = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_StackSeries')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_StackSeries bit NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_StackSeries = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_ShowValues')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_ShowValues bit NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_ShowValues = 0'
	END
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_viewID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_viewID int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_viewID = 0'
	END
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_TableID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_TableID int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_TableID = 0'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_ColumnID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_ColumnID int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_ColumnID = 0'
	END	
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_FilterID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_FilterID int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_FilterID = 0'
	END		
	
	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Chart_AggregateType')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Chart_AggregateType int NULL'
		EXEC sp_executesql N'UPDATE ASRSysSSIntranetLinks SET Chart_AggregateType = 0'
	END		

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIViews', 'U') AND name = 'WFOutOfOffice')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIViews ADD WFOutOfOffice bit NOT NULL DEFAULT 1'
	END		

--UPDATE EXISTING SEPARATORS 

	IF NOT EXISTS(SELECT id FROM syscolumns
	              WHERE  id = OBJECT_ID('ASRSysSSIntranetLinks', 'U') AND name = 'Element_Type')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysSSIntranetLinks ADD Element_Type int NULL'
		EXEC sp_executesql N'UPDATE [ASRSysSSIntranetLinks] SET Element_Type = convert(int, [ASRSysSSIntranetLinks].IsSeparator)'
		EXEC sp_executesql N'ALTER TABLE dbo.ASRSysSSIntranetLinks DROP COLUMN IsSeparator'

		EXEC sp_executesql N'UPDATE [ASRSysSSIntranetLinks] SET [text] = [prompt] WHERE [text] = ''<SEPARATOR>'''
		EXEC sp_executesql N'UPDATE [ASRSysSSIntranetLinks] SET [prompt] = ''<SEPARATOR>'' WHERE [Element_Type] = 1'
	END		

	
--Create New SSI Dashboard Stored Procedures

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRIntGetPicture]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRIntGetPicture];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRIntGetPicture]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRIntGetPicture]	
	(
		@piPictureID		integer
	)
	AS
	BEGIN
		SET NOCOUNT ON
		SELECT TOP 1 name, picture
		FROM ASRSysPictures
		WHERE pictureID = @piPictureID
	END';

	EXECUTE sp_executeSQL @sSPCode;



	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRIntShowOutOfOfficeHyperlink]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRIntShowOutOfOfficeHyperlink];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRIntShowOutOfOfficeHyperlink]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRIntShowOutOfOfficeHyperlink]	
	(
		@piTableID		integer,
		@piViewID		integer,
		@pfDisplayHyperlink	bit 	OUTPUT
	)
	AS
	BEGIN
		SELECT @pfDisplayHyperlink = WFOutOfOffice
		FROM ASRSysSSIViews
		WHERE (TableID = @piTableID) 
			AND  (ViewID = @piViewID)
	END';

	EXECUTE sp_executeSQL @sSPCode;


	----update the SPASRINTGETLINKS stored procedure

	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRIntGetLinks]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRIntGetLinks];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRIntGetLinks]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRIntGetLinks] 
(
		@plngTableID	integer,
		@plngViewID		integer
)
AS
BEGIN
	DECLARE @iCount			integer,
		@iUtilType			integer, 
		@iUtilID			integer,
		@iScreenID			integer,
		@iTableID			integer,
		@sTableName			sysname,
		@iTableType			integer,
		@sRealSource		sysname,
		@iChildViewID		integer,
		@sAccess			varchar(MAX),
		@sGroupName			varchar(255),
		@pfPermitted		bit,
		@sActualUserName	sysname,
		@iActualUserGroupID integer,
		@fBaseTableReadable bit,
		@iBaseTableID		integer,
		@sURL				varchar(MAX), 
		@fUtilOK			bit,
		@iLinkType			integer;		-- 0 = Hypertext, 1 = Button, 2 = Dropdown List
	
	SET NOCOUNT ON;
	IF @plngViewID < 1 
	BEGIN 
		SET @plngViewID = -1;
	END
	SET @fBaseTableReadable = 1;
	
	IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) <> ''SA''
	BEGIN
		EXEC [dbo].[spASRIntGetActualUserDetails]
			@sActualUserName OUTPUT,
			@sGroupName OUTPUT,
			@iActualUserGroupID OUTPUT;
		
		DECLARE @Phase1 TABLE([ID] INT);
		INSERT INTO @Phase1
			SELECT Object_ID(ASRSysViews.ViewName) 
			FROM ASRSysViews 
			WHERE NOT Object_ID(ASRSysViews.ViewName) IS null
			UNION
			SELECT Object_ID(ASRSysTables.TableName) 
			FROM ASRSysTables 
			WHERE NOT Object_ID(ASRSysTables.TableName) IS null
			UNION
			SELECT OBJECT_ID(left(''ASRSysCV'' + convert(varchar(1000), ASRSysChildViews2.childViewID) 
				+ ''#'' + replace(ASRSysTables.tableName, '' '', ''_'')
				+ ''#'' + replace(@sGroupName, '' '', ''_''), 255))
			FROM ASRSysChildViews2
			INNER JOIN ASRSysTables 
				ON ASRSysChildViews2.tableID = ASRSysTables.tableID
			WHERE NOT OBJECT_ID(left(''ASRSysCV'' + convert(varchar(1000), ASRSysChildViews2.childViewID) 
				+ ''#'' + replace(ASRSysTables.tableName, '' '', ''_'')
				+ ''#'' + replace(@sGroupName, '' '', ''_''), 255)) IS null;
		-- Cached view of the sysprotects table
		DECLARE @SysProtects TABLE([ID] int PRIMARY KEY CLUSTERED);
		INSERT INTO @SysProtects
			SELECT p.[ID] 
			FROM #sysprotects p
						INNER JOIN SysColumns c ON (c.id = p.id
							AND c.[Name] = ''timestamp''
							AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
							AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) != 0)
							OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
							AND (convert(int,substring(p.columns,c.colid/8+1,1))&power(2,c.colid&7)) = 0)))
			WHERE p.[ProtectType] IN (204, 205)
				AND p.[Action] = 193			
				AND p.id IN (SELECT ID FROM @Phase1);
		-- Readable tables
		DECLARE @ReadableTables TABLE([Name] sysname PRIMARY KEY CLUSTERED);
		INSERT INTO @ReadableTables
			SELECT OBJECT_NAME(p.ID)
			FROM @SysProtects p;
		
		SET @sRealSource = '''';
		IF @plngViewID > 0
		BEGIN
			SELECT @sRealSource = viewName
			FROM ASRSysViews
			WHERE viewID = @plngViewID;
		END
		ELSE
		BEGIN
			SELECT @sRealSource = tableName
			FROM ASRSysTables
			WHERE tableID = @plngTableID;
		END
		SET @fBaseTableReadable = 0
		IF len(@sRealSource) > 0
		BEGIN
			SELECT @iCount = COUNT(*)
			FROM @ReadableTables
			WHERE name = @sRealSource;
		
			IF @iCount > 0
			BEGIN
				SET @fBaseTableReadable = 1;
			END
		END
	END
	DECLARE @Links TABLE([ID]						integer PRIMARY KEY CLUSTERED,
											 [utilityType]	integer,
											 [utilityID]		integer,
											 [screenID]			integer,
											 [LinkType]			integer);
	INSERT INTO @Links ([ID],[utilityType],[utilityID],[screenID], [LinkType])
	SELECT ASRSysSSIntranetLinks.ID,
					ASRSysSSIntranetLinks.utilityType,
					ASRSysSSIntranetLinks.utilityID,
					ASRSysSSIntranetLinks.screenID,
					ASRSysSSIntranetLinks.LinkType
	FROM ASRSysSSIntranetLinks
	WHERE (viewID = @plngViewID
			AND tableid = @plngTableID)
			AND (id NOT IN (SELECT linkid 
								FROM ASRSysSSIHiddenGroups
								WHERE groupName = @sGroupName));
	/* Remove any utility links from the temp table where the utility has been deleted or hidden from the current user.*/
	/* Or if the user does not permission to run them. */	
	DECLARE utilitiesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysSSIntranetLinks.utilityType,
					ASRSysSSIntranetLinks.utilityID,
					ASRSysSSIntranetLinks.screenID,
					ASRSysSSIntranetLinks.LinkType
	FROM ASRSysSSIntranetLinks
	WHERE (viewID = @plngViewID
				AND tableid = @plngTableID)
			AND (utilityID > 0 
				OR screenID > 0);
	OPEN utilitiesCursor;
	FETCH NEXT FROM utilitiesCursor INTO @iUtilType, @iUtilID, @iScreenID, @iLinkType;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iUtilID > 0
		BEGIN
			SET @fUtilOK = 1	;			
			/* Check if the utility is deleted or hidden from the user. */
			EXECUTE [dbo].[spASRIntCurrentAccessForRole]
								@sGroupName,
								@iUtilType,
								@iUtilID,
								@sAccess	OUTPUT;
			IF @sAccess = ''HD'' 
			BEGIN
				/* Report/utility is hidden from the user. */
				SET @fUtilOK = 0;
			END
			IF @fUtilOK = 1
			BEGIN
				/* Check if the user has system permission to run this type of report/utility. */
				IF UPPER(LTRIM(RTRIM(SYSTEM_USER))) <> ''SA''
				BEGIN
					SELECT @pfPermitted = ASRSysGroupPermissions.permitted
					FROM ASRSysPermissionItems
					INNER JOIN ASRSysPermissionCategories 
					ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID
						AND ASRSysPermissionCategories.categoryKey = 					
							CASE 
								WHEN @iUtilType = 17 THEN ''CALENDARREPORTS''
								WHEN @iUtilType = 9 THEN ''MAILMERGE''
								WHEN @iUtilType = 2 THEN ''CUSTOMREPORTS''
								WHEN @iUtilType = 25 THEN ''WORKFLOW''
								ELSE ''''
							END
					LEFT OUTER JOIN ASRSysGroupPermissions 
					ON ASRSysPermissionItems.itemID = ASRSysGroupPermissions.itemID
						AND ASRSysGroupPermissions.groupName = @sGroupName
					WHERE ASRSysPermissionItems.itemKey = ''RUN'';
					IF (@pfPermitted IS null) OR (@pfPermitted = 0)
					BEGIN
						/* User does not have system permission to run this type of report/utility. */
						SET @fUtilOK = 0;
					END
				END
			END
			IF @fUtilOK = 1
			BEGIN
				/* Check if the user has read permission on the report/utility base table or any views on it. */
				SET @iBaseTableID = 0;
				IF @iUtilType = 17 /* Calendar Reports */
				BEGIN
					SELECT @iBaseTableID = baseTable
					FROM ASRSysCalendarReports
					WHERE id = @iUtilID;
				END
				IF @iUtilType = 2 /* Custom Reports */
				BEGIN
					SELECT @iBaseTableID = baseTable
					FROM ASRSysCustomReportsName
					WHERE id = @iUtilID;
				END
				IF @iUtilType = 9 /* Mail Merge */
				BEGIN
					SELECT @iBaseTableID = TableID
					FROM ASRSysMailMergeName
					WHERE MailMergeID = @iUtilID;
				END
				/* Not check required for reports/utilities without a base table.
				OR reports/utilities based on the top-level table if the user has read permission on the current view. */
				IF (@iBaseTableID > 0)
					AND((@fBaseTableReadable = 0)
						OR (@iBaseTableID <> @plngTableID))
				BEGIN
					IF (@iLinkType <> 0) -- Hypertext link
						AND (@fBaseTableReadable = 0)
						AND (@iBaseTableID = @plngTableID)
					BEGIN
						/* The report/utility is based on the top-level table, and the user does NOT have read permission
						on the current view (on which Button & DropdownList links are scoped). */
						SET @fUtilOK = 0;
					END
					ELSE
					BEGIN
						SELECT @iCount = COUNT(p.ID)
						FROM @SysProtects p
						WHERE OBJECT_NAME(p.ID) IN (SELECT ASRSysTables.tableName
							FROM ASRSysTables
							WHERE ASRSysTables.tableID = @iBaseTableID
						UNION 
							SELECT ASRSysViews.viewName
								FROM ASRSysViews
								WHERE ASRSysViews.viewTableID = @iBaseTableID
						UNION
							SELECT
								left(''ASRSysCV'' 
									+ convert(varchar(1000), ASRSysChildViews2.childViewID) 
									+ ''#''
									+ replace(ASRSysTables.tableName, '' '', ''_'')
									+ ''#''
									+ replace(@sGroupName, '' '', ''_''), 255)
								FROM ASRSysChildViews2
								INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID
								WHERE ASRSysChildViews2.role = @sGroupName
									AND ASRSysChildViews2.tableID = @iBaseTableID);
						IF @iCount = 0 
						BEGIN
							SET @fUtilOK = 0;
						END
					END
				END
			END
			/* For some reason the user cannot use this report/utility, so remove it from the temp table of links. */
			IF @fUtilOK = 0 
			BEGIN
				DELETE FROM @Links
				WHERE utilityType = @iUtilType
					AND utilityID = @iUtilID;
			END
		END
		
		IF (@iScreenID > 0) AND (UPPER(LTRIM(RTRIM(SYSTEM_USER))) <> ''SA'')
		BEGIN
			/* Do not display the link if the user does not have permission to read the defined view/table for the screen. */
			SELECT @iTableID = ASRSysTables.tableID, 
				@sTableName = ASRSysTables.tableName,
				@iTableType = ASRSysTables.tableType
			FROM ASRSysScreens
						INNER JOIN ASRSysTables 
						ON ASRSysScreens.tableID = ASRSysTables.tableID
			WHERE screenID = @iScreenID;
			SET @sRealSource = '''';
			IF @iTableType  = 2
			BEGIN
				SET @iChildViewID = 0;
				/* Child table - check child views. */
				SELECT @iChildViewID = childViewID
				FROM ASRSysChildViews2
				WHERE tableID = @iTableID
					AND [role] = @sGroupName;
				
				IF @iChildViewID IS null SET @iChildViewID = 0;
				
				IF (@iChildViewID > 0) AND (@fBaseTableReadable = 1)
				BEGIN
					SET @sRealSource = ''ASRSysCV'' + 
						convert(varchar(1000), @iChildViewID) +
						''#'' + replace(@sTableName, '' '', ''_'') +
						''#'' + replace(@sGroupName, '' '', ''_'');
				
					SET @sRealSource = left(@sRealSource, 255);
				END
				ELSE
				BEGIN
					DELETE FROM @Links
					WHERE screenID = @iScreenID;
				END
			END
			ELSE
			BEGIN
				/* Not a child table - must be the top-level table. Check if the user has ''read'' permission on the defined view. */
				SET @sRealSource = '''';
				IF @plngViewID > 0
				BEGIN
					SELECT @sRealSource = viewName
					FROM ASRSysViews
					WHERE viewID = @plngViewID;
				END
				ELSE
				BEGIN
					SELECT @sRealSource = tableName
					FROM ASRSysTables
					WHERE tableID = @plngTableID;
				END
			END
			IF len(@sRealSource) > 0
			BEGIN
				SELECT @iCount = COUNT(*)
				FROM @ReadableTables
				WHERE name = @sRealSource;
			
				IF @iCount = 0
				BEGIN
					DELETE FROM @Links
					WHERE screenID = @iScreenID;
				END
			END
		END
		FETCH NEXT FROM utilitiesCursor INTO @iUtilType, @iUtilID, @iScreenID, @iLinkType;
	END
	CLOSE utilitiesCursor;
	DEALLOCATE utilitiesCursor;
	/* Remove the Workflow links if the URL has not been configured. */
	SELECT @sURL = isnull(settingValue , '''')
	FROM ASRSysSystemSettings
	WHERE section = ''MODULE_WORKFLOW''		
		AND settingKey = ''Param_URL'';
	
	IF LEN(@sURL) = 0
	BEGIN
		DELETE FROM @Links
		WHERE utilityType = 25;
	END
	SELECT ASRSysSSIntranetLinks.*, 
		CASE 
			WHEN ASRSysSSIntranetLinks.utilityType = 9 THEN ASRSysMailMergeName.TableID
			WHEN ASRSysSSIntranetLinks.utilityType = 2 THEN ASRSysCustomReportsName.baseTable
			WHEN ASRSysSSIntranetLinks.utilityType = 17 THEN ASRSysCalendarReports.baseTable
			WHEN ASRSysSSIntranetLinks.utilityType = 25 THEN 0
			ELSE null
		END AS [baseTable],
		ASRSysColumns.ColumnName as [Chart_ColumnName]
	FROM ASRSysSSIntranetLinks
			LEFT OUTER JOIN ASRSysMailMergeName 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysMailMergeName.MailMergeID
				AND ASRSysSSIntranetLinks.utilityType = 9
			LEFT OUTER JOIN ASRSysCalendarReports 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysCalendarReports.ID
				AND ASRSysSSIntranetLinks.utilityType = 17
			LEFT OUTER JOIN ASRSysCustomReportsName 
			ON ASRSysSSIntranetLinks.utilityID = ASRSysCustomReportsName.ID
				AND ASRSysSSIntranetLinks.utilityType = 2
			LEFT OUTER JOIN ASRSysColumns
			ON ASRSysSSIntranetLinks.Chart_ColumnID = ASRSysColumns.columnID
	WHERE ASRSysSSIntranetLinks.ID IN (SELECT ID FROM @Links)
	ORDER BY ASRSysSSIntranetLinks.linkOrder;
END';

	EXECUTE sp_executeSQL @sSPCode;


---Create the new spASRIntGetSSIWelcomeDetails stored procedure 


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[spASRIntGetSSIWelcomeDetails]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[spASRIntGetSSIWelcomeDetails];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[spASRIntGetSSIWelcomeDetails]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[spASRIntGetSSIWelcomeDetails]  
(
            @piWelcomeColumnID integer,
            @piSingleRecordViewID integer,
            @psUserName varchar(255),     
            @psWelcomeMessage varchar(255) OUTPUT
)
AS
BEGIN

      DECLARE @sql nvarchar(max)
      DECLARE @dtLastLogon datetime
      DECLARE @myval varchar(max)
      DECLARE @psLogonTime varchar(20)
      DECLARE @psLogonDay varchar(20)
      DECLARE @psWelcomeName varchar(255)
      DECLARE @psLastLogon varchar(50)          

      --- try to get the users welcome name

      BEGIN TRY
            SELECT @sql = ''SELECT @outparm = [''+c.columnname+''] FROM [''+v.viewname+'']''
                  FROM ASRSysColumns c, ASRSysViews v
                  WHERE c.columnID = @piWelcomeColumnID AND v.ViewID = @piSingleRecordViewID

            EXEC sp_executesql @sql, N''@outparm nvarchar(max) output'', @myval OUTPUT
      
            IF LEN(LTRIM(RTRIM(@myval))) = 0 OR @@ROWCOUNT = 0 or ISNULL(@myval, '''') = ''''
            BEGIN
                  SET @psWelcomeName = ''''
            END
            ELSE
            BEGIN
                  SET @psWelcomeName = '' '' + isnull(@myval, '''')
            END

      END TRY
      
      BEGIN CATCH
            SET @psWelcomeName = ''''
      END CATCH
      
      --- Now get the last logon details

      SELECT top 2 @dtLastLogon = DateTimeStamp
            FROM ASRSysAuditAccess WHERE [UserName] = @psUserName
            AND [HRProModule] = ''Intranet'' AND [Action] = ''log in''
      ORDER BY DateTimeStamp DESC

      IF @@ROWCOUNT > 0 
      BEGIN
            SET @psLogonTime = CONVERT(varchar(5), @dtLastLogon, 108)
            SELECT @psLogonDay = 
                  CASE datediff(day, @dtLastLogon, GETDATE())
                  WHEN 0 THEN ''today''
                  WHEN 1 THEN ''yesterday''
                  ELSE ''on '' + CAST(DAY(@dtLastLogon) AS VARCHAR(2)) + '' '' + DATENAME(MM, @dtLastLogon) + '' '' + CAST(YEAR(@dtLastLogon) AS VARCHAR(4))
            END
            SET @psWelcomeMessage = ''Welcome back'' + @psWelcomeName + '', you last logged in at '' + @psLogonTime + '' '' + @psLogonDay
      END
      ELSE
      BEGIN
            SET @psWelcomeMessage = ''Welcome '' + @psWelcomeName
      END

END';

	EXECUTE sp_executeSQL @sSPCode;


/* ------------------------------------------------------------- */
PRINT 'Step 6 - Shared Table Integration'


	-- Add columns to ASRSysMailMergeName
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysAccordTransactions', 'U') AND name = 'BatchID')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE dbo.[ASRSysAccordTransactions]
								 ADD [BatchID] integer';
	END


/* ------------------------------------------------------------- */
PRINT 'Step 7 - Validation Messages'

	EXECUTE spASRResizeColumn 'ASRSysColumns','errorMessage','MAX';



/* ------------------------------------------------------------- */
PRINT 'Step 8 - Updating Support Contact Details'
/* ------------------------------------------------------------- */

	---New email address
  SET @sSPCode_0 = 'DELETE FROM [ASRSysSystemSettings] 
              WHERE [ASRSysSystemSettings].[Section] = ''support'' 
                AND [ASRSysSystemSettings].[SettingKey] = ''email'''
	EXECUTE (@sSPCode_0)
	
  SET @sSPCode_0 = 'INSERT INTO [ASRSysSystemSettings] ([ASRSysSystemSettings].[Section], 
                                                  [ASRSysSystemSettings].[SettingKey], 
                                                  [ASRSysSystemSettings].[SettingValue])
                    VALUES (''support'', ''email'', ''service.delivery@coasolutions.com'')'
              
	EXECUTE (@sSPCode_0)

	---New Web access detail
  SET @sSPCode_0 = 'DELETE FROM [ASRSysSystemSettings] 
              WHERE [ASRSysSystemSettings].[Section] = ''support'' 
                AND [ASRSysSystemSettings].[SettingKey] = ''webpage'''
	EXECUTE (@sSPCode_0)
	
  SET @sSPCode_0 = 'INSERT INTO [ASRSysSystemSettings] ([ASRSysSystemSettings].[Section], 
                                                  [ASRSysSystemSettings].[SettingKey], 
                                                  [ASRSysSystemSettings].[SettingValue])
                    VALUES (''support'', ''webpage'', ''http://webfirst.coasolutions.com'')'
              
	EXECUTE (@sSPCode_0)


	---New contact number (from UK)
  SET @sSPCode_0 = 'DELETE FROM [ASRSysSystemSettings] 
              WHERE [ASRSysSystemSettings].[Section] = ''support'' 
                AND [ASRSysSystemSettings].[SettingKey] = ''telephone no'''
	EXECUTE (@sSPCode_0)
	
  SET @sSPCode_0 = 'INSERT INTO [ASRSysSystemSettings] ([ASRSysSystemSettings].[Section], 
                                                  [ASRSysSystemSettings].[SettingKey], 
                                                  [ASRSysSystemSettings].[SettingValue])
                    VALUES (''support'', ''telephone no'', ''08451 609 999'')'
              
	EXECUTE (@sSPCode_0)

	---New contact number (International)
  SET @sSPCode_0 = 'DELETE FROM [ASRSysSystemSettings] 
              WHERE [ASRSysSystemSettings].[Section] = ''support'' 
                AND [ASRSysSystemSettings].[SettingKey] = ''telephone no intl'''
	EXECUTE (@sSPCode_0)
	
  SET @sSPCode_0 = 'INSERT INTO [ASRSysSystemSettings] ([ASRSysSystemSettings].[Section], 
                                                  [ASRSysSystemSettings].[SettingKey], 
                                                  [ASRSysSystemSettings].[SettingValue])
                    VALUES (''support'', ''telephone no intl'', ''+44(0)1932 590 721'')'
              
	EXECUTE (@sSPCode_0)
	


/* ------------------------------------------------------------- */
PRINT 'Step 9 - Misc stored procedures'
/* ------------------------------------------------------------- */

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
					WHERE ColumnID = @ColumnID And @RecordID = RecordID
					ORDER BY DateTimeStamp DESC);
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
		
				SELECT @psResult = ISNULL(settingValue , '''')
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
		
				SELECT @psResult = ISNULL(settingValue , '''')
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
					IF UPPER(@Char) = ''M'' AND UPPER(SUBSTRING(@psInput, @Index + 1, 1)) = ''C'' AND UPPER(SUBSTRING(@psInput, @Index - 1, 1)) = ''''
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
					SET @psResult = RTRIM(LTRIM(LEFT(@psForenames, CHARINDEX('' '', @psForenames))));
				END
				ELSE
				BEGIN
					SET @psResult = RTRIM(LTRIM(@psForenames));
				END
			END
		END';

	EXECUTE sp_executeSQL @sSPCode;
	
	
PRINT 'Step 10 - New Shared Table Transfer Types'

	-- Leave Reason
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 101
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (101, ''Leave Reason'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,101,1,''Code Table ID'',0,0,2,1,1,2,''11'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,101,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,101,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,101,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,101,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,101,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,101,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,101,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,101,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,101,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,101,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,101,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,101,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,101,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,101,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END
	
		-- Marital Status
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 102
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (102, ''Marital Status'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMApType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,102,1,''Code Table ID'',0,0,2,1,1,2,''10'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,102,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,102,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,102,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,102,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,102,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,102,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,102,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,102,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,102,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,102,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,102,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,102,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,102,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,102,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Cost Centre
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 103
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (103, ''Cost Centre'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMApType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,103,1,''Code Table ID'',0,0,2,1,1,2,''2'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,103,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,103,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,103,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,103,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,103,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,103,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,103,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,103,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,103,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,103,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,103,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,103,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,103,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,103,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,103,0,''Project'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Job Grade
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 104
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (104, ''Job Grade'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,104,1,''Code Table ID'',0,0,2,1,1,2,''3'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,104,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,104,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,104,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,104,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,104,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,104,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,104,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,104,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,104,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,104,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,104,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,104,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,104,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,104,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Sort Code 1
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 105
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (105, ''Sort Code 1'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,105,1,''Code Table ID'',0,0,2,1,1,2,''4'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,105,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,105,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,105,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,105,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,105,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,105,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,105,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,105,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,105,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,105,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,105,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,105,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,105,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,105,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Sort Code 2
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 106
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (106, ''Sort Code 2'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,106,1,''Code Table ID'',0,0,2,1,1,2,''5'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,106,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,106,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,106,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,106,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,106,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,106,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,106,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,106,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,106,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,106,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,106,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,106,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,106,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,106,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END
	
	-- Sort Code 3
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 107
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (107, ''Sort Code 3'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,107,1,''Code Table ID'',0,0,2,1,1,2,''6'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,107,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,107,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,107,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,107,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,107,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,107,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,107,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,107,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,107,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,107,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,107,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,107,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,107,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,107,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Sort Code 4
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 108
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (108, ''Sort Code 4'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,108,1,''Code Table ID'',0,0,2,1,1,2,''7'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,108,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,108,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,108,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,108,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,108,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,108,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,108,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,108,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,108,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,108,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,108,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,108,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,108,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,108,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Sort Code 5
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 109
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (109, ''Sort Code 5'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,109,1,''Code Table ID'',0,0,2,1,1,2,''8'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,109,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,109,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,109,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,109,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,109,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,109,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,109,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,109,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,109,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,109,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,109,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,109,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,109,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,109,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Sort Code 6
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 110
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (110, ''Sort Code 6'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,110,1,''Code Table ID'',0,0,2,1,1,2,''9'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,110,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,110,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,110,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,110,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,110,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,110,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,110,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,110,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,110,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,110,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,110,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,110,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,110,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,110,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Job Title
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 111
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (111, ''Job Title'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,111,1,''Code Table ID'',0,0,2,1,1,2,''14'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,111,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,111,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,111,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,111,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,111,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,111,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,111,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,111,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,111,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,111,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,111,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,111,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,111,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,111,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Ethnic Origin
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 112
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (112, ''Ethnic Origin'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,112,1,''Code Table ID'',0,0,2,1,1,2,''13'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,112,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,112,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,112,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,112,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,112,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,112,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,112,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,112,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,112,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,112,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,112,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,112,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,112,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,112,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Nationality
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 113
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (113, ''Nationality'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,113,1,''Code Table ID'',0,0,2,1,1,2,''12'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,113,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,113,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,113,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,113,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,113,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,113,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,113,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,113,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,113,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,113,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,113,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,113,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,113,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,113,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Reports To (1)
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 114
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (114, ''Reports To (1)'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,114,1,''Code Table ID'',0,0,2,1,1,2,''15'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,114,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,114,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,114,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,114,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,114,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,114,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,114,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,114,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,114,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,114,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,114,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,114,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,114,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,114,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Reports To (2)
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 115
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (115, ''Reports To (2)'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,115,1,''Code Table ID'',0,0,2,1,1,2,''16'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,115,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,115,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,115,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,115,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,115,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,115,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,115,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,115,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,115,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,115,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,115,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,115,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,115,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,115,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Absence Type
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 116
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (116, ''Absence Type'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,116,1,''Code Table ID'',0,0,2,1,1,2,''900'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,116,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,116,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,116,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,116,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,116,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,116,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,116,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,116,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,116,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,116,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,116,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,116,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,116,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,116,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,116,0,''OSP Indicator'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (16,116,0,''SSP Indicator'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (17,116,0,''Days/Hours'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Absence Reason
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 117
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (117, ''Absence Reason'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,117,1,''Code Table ID'',0,0,2,1,1,2,''901'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,117,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,117,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,117,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,117,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,117,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,117,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,117,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,117,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,117,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,117,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,117,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,117,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,117,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,117,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (15,117,0,''Absence Type'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Bank Details
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 118
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (118, ''Bank Details'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (0,118,1,''Sort Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,118,1,''Bank Name'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,118,0,''Branch'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,118,0,''Address 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,118,0,''Address 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,118,0,''Address 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,118,0,''Address 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,118,0,''Address 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Extra Code Table - User Defined 1
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 131
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (131, ''Extra Code Table - User Defined 1'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,131,1,''Code Table ID'',0,0,2,1,1,2,''100'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,131,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,131,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,131,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,131,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,131,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,131,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,131,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,131,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,131,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,131,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,131,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,131,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,131,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,131,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Extra Code Table - User Defined 2
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 132
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (132, ''Extra Code Table - User Defined 2'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,132,1,''Code Table ID'',0,0,2,1,1,2,''101'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,132,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,132,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,132,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,132,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,132,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,132,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,132,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,132,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,132,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,132,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,132,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,132,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,132,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,132,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Extra Code Table - User Defined 3
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 133
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (133, ''Extra Code Table - User Defined 3'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,133,1,''Code Table ID'',0,0,2,1,1,2,''102'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,133,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,133,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,133,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,133,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,133,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,133,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,133,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,133,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,133,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,133,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,133,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,133,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,133,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,133,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Extra Code Table - User Defined 4
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 134
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (134, ''Extra Code Table - User Defined 4'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,134,1,''Code Table ID'',0,0,2,1,1,2,''103'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,134,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,134,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,134,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,134,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,134,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,134,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,134,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,134,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,134,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,134,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,134,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,134,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,134,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,134,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Extra Code Table - User Defined 5
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 135
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (135, ''Extra Code Table - User Defined 5'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,135,1,''Code Table ID'',0,0,2,1,1,2,''104'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,135,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,135,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,135,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,135,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,135,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,135,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,135,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,135,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,135,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,135,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,135,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,135,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,135,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,135,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Extra Code Table - User Defined 6
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 136
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (136, ''Extra Code Table - User Defined 6'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,136,1,''Code Table ID'',0,0,2,1,1,2,''105'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,136,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,136,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,136,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,136,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,136,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,136,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,136,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,136,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,136,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,136,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,136,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,136,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,136,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,136,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Extra Code Table - User Defined 7
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 137
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (137, ''Extra Code Table - User Defined 7'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,137,1,''Code Table ID'',0,0,2,1,1,2,''106'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,137,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,137,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,137,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,137,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,137,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,137,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,137,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,137,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,137,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,137,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,137,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,137,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,137,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,137,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Extra Code Table - User Defined 8
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 138
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (138, ''Extra Code Table - User Defined 8'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,138,1,''Code Table ID'',0,0,2,1,1,2,''107'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,138,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,138,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,138,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,138,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,138,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,138,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,138,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,138,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,138,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,138,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,138,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,138,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,138,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,138,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END

	-- Extra Code Table - User Defined 9
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 139
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (139, ''Extra Code Table - User Defined 9'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,139,1,''Code Table ID'',0,0,2,1,1,2,''108'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,139,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,139,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,139,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,139,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,139,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,139,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,139,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,139,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,139,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,139,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,139,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,139,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,139,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,139,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END
	
	-- Extra Code Table - User Defined 10
	SELECT @iRecCount = count(TransferTypeID) FROM ASRSysAccordTransferTypes WHERE TransferTypeID = 140
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferTypes  (TransferTypeID, TransferType, ASRBaseTableID, FilterID, ForceAsUpdate, IsVisible) VALUES (140, ''Extra Code Table - User Defined 10'' ,0,0,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer, ASRMapType, ASRValue, ASRColumnID, ASRExprID) VALUES (0,140,1,''Code Table ID'',0,0,2,1,1,2,''109'',0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (1,140,1,''Code'',0,0,2,1,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (2,140,1,''Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (3,140,1,''Short Description'',0,0,2,0,1)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (4,140,0,''Email Address'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (5,140,0,''Supplementary Field 1a'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (6,140,0,''Supplementary Field 1b'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (7,140,0,''Supplementary Field 1c'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (8,140,0,''Supplementary Field 1d'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (9,140,0,''Supplementary Field 1e'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (10,140,0,''Supplementary Field 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (11,140,0,''Supplementary Field 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,140,0,''Supplementary Field 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (13,140,0,''Supplementary Field 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (14,140,0,''Supplementary Field 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END
	
	-- Updates to Employee Shared Table
	SELECT @iRecCount = count(TransferFieldID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferFieldID = 195 AND TransferTypeID = 0
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (195,0,0,''Analysis Code 1'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (196,0,0,''Analysis Code 2'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (197,0,0,''Analysis Code 3'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (198,0,0,''Analysis Code 4'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (199,0,0,''Analysis Code 5'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (200,0,0,''Analysis Code 6'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand
	
	END
	
	-- Update to Absence Shared Table
	SELECT @iRecCount = count(TransferFieldID) FROM ASRSysAccordTransferFieldDefinitions WHERE TransferFieldID = 12 AND TransferTypeID = 72
	IF @iRecCount = 0
	BEGIN

		SELECT @NVarCommand = 'INSERT INTO ASRSysAccordTransferFieldDefinitions  (TransferFieldID, TransferTypeID, Mandatory, Description, IsCompanyCode, IsEmployeeCode, Direction, IsKeyField, AlwaysTransfer) VALUES (12,72,0,''Memo'',0,0,2,0,0)'
		EXEC sp_executesql @NVarCommand

	END


/* ------------------------------------------------------------- */
PRINT 'Step 11 - Update Round To Nearest Number Function'


	IF EXISTS (SELECT *
		FROM dbo.sysobjects
		WHERE id = object_id(N'[dbo].[sp_ASRFn_RoundToNearestNumber]')
			AND OBJECTPROPERTY(id, N'IsProcedure') = 1)
		DROP PROCEDURE [dbo].[sp_ASRFn_RoundToNearestNumber];

	SET @sSPCode = 'CREATE PROCEDURE [dbo].[sp_ASRFn_RoundToNearestNumber]
		AS
		BEGIN
			DECLARE @iDummy integer;
		END';
	EXECUTE sp_executeSQL @sSPCode;

	SET @sSPCode = 'ALTER PROCEDURE [dbo].[sp_ASRFn_RoundToNearestNumber]
		(
			@pfReturn 			float OUTPUT,
			@pfNumberToRound 	float,
			@pfNearestNumber	float
		)
		AS
		BEGIN

			DECLARE @pfRemainder float;

			/* Calculate the remainder. Cannot use the % because it only works on integers and not floats. */
			set @pfReturn = 0;
			if @pfNearestNumber <= 0 return
	
			set @pfRemainder = @pfNumberToRound - (floor(@pfNumberToRound / @pfNearestNumber) * @pfNearestNumber);

			/* Formula for rounding to the nearest specified number */
			if ((@pfNumberToRound < 0) AND (@pfRemainder <= (@pfNearestNumber / 2.0)))
				OR ((@pfNumberToRound >= 0) AND (@pfRemainder < (@pfNearestNumber / 2.0)))
					set @pfReturn = @pfNumberToRound - @pfRemainder;
				else
					set @pfReturn = @pfNumberToRound + @pfNearestNumber - @pfRemainder;

		END';

	EXECUTE sp_executeSQL @sSPCode;


	
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
values('database', 'version', '4.1')

delete from asrsyssystemsettings
where [Section] = 'intranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('intranet', 'minimum version', '4.1.0')

delete from asrsyssystemsettings
where [Section] = 'ssintranet' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('ssintranet', 'minimum version', '4.1.0')

delete from asrsyssystemsettings
where [Section] = 'server dll' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('server dll', 'minimum version', '3.4.0')

delete from asrsyssystemsettings
where [Section] = '.NET Assembly' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('.NET Assembly', 'minimum version', '4.1.0')

delete from asrsyssystemsettings
where [Section] = 'outlook service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('outlook service', 'minimum version', '4.1.0')

delete from asrsyssystemsettings
where [Section] = 'workflow service' and [SettingKey] = 'minimum version'
insert ASRSysSystemSettings([Section], [SettingKey], [SettingValue])
values('workflow service', 'minimum version', '4.1.0')

insert into asrsysauditaccess
(DateTimeStamp, UserGroup, UserName, ComputerName, HRProModule, Action)
values (getdate(),'<none>',left(system_user,50),lower(left(host_name(),30)),'System','v4.1')


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
PRINT 'Update Script Has Converted Your HR Pro Database To Use v4.1 Of HR Pro'
