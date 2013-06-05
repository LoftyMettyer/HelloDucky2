
/* --------------------------------------------------- */
/* Update the database from version 4.0 to version 4.1 */
/* --------------------------------------------------- */

DECLARE @iRecCount integer,
	@sDBVersion varchar(10),
	@DBName varchar(255),
	@Command varchar(max),
	@iSQLVersion int,
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

-- Only allow script to be run on SQL2005 or above
SELECT @iSQLVersion = convert(float,substring(@@version,charindex('-',@@version)+2,2))
IF (@iSQLVersion <> 9 AND @iSQLVersion <> 10)
BEGIN
	RAISERROR('The SQL Server is incompatible with this version of HR Pro', 16, 1)
	RETURN
END


/* ------------------------------------------------------------- */
PRINT 'Step 1 - Modifying Workflow procedures'

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


	-- Add columns to ASRSysMailMergeName
	IF NOT EXISTS(SELECT id FROM syscolumns WHERE  id = OBJECT_ID('ASRSysMailMergeName', 'U') AND name = 'OutputPrinterName')
    BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysMailMergeName
								 ADD [OutputPrinterName] nvarchar(255), [DocumentMapID] integer';
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

		WRITETEXT ASRSysPermissionCategories.picture @ptrval 0x000001000100101000000100080068050000160000002800000010000000200000000100080000000000000000000000000000000000000000000000000000000000070306000F060D000E0D0E00100F0F001411130024162000281523002A1B26003B162D003D172F002C2C2C0032282E00302F2F00372131003B3B3A00401F3200491A360061224E007329540072396100404040004F4F4E005648510051505000575655005E5E5E007C496F00606060006766650077767600787776007E737A0081376D008F467A008D507C00964F82009A5586009F5E8D0090648300A8759900AD759D00AB7F9900448BF000468CF000488DF0004F92F0005093F1005295F1005798F2005E9CF30070A7F30071A9F50074ABF50076ACF500918F8E00979796009D9C9A009E9E9E00A9829C00A09F9D00B684A700B68DA500B786A900B887A900BB8CAD00B799AD00A2A2A100A8A6A500A8A8A800B3B2B100C195B300C297B400C399B600C59BB700C7A2BB00C8A1BC00CAA5BE00B6B3D90080AFF20080B1F5008BB9F70092BEF800D5B5C900D4B8CA00D7B9CC00B0D1FA00D1D0CF00D4D3D300D7D6D500D9D8D800DFDEDE00E1D2DC00E0DFDE00E5E4E400EFE1E900EBEAEA00ECEBEB00F5F5F500FBFAFA00FDFDFD004CB0000059CF000067F0000078FF11008AFF31009CFF5100AEFF7100C0FF9100D2FFB100E4FFD100FFFFFF0000000000262F0000405000005A700000749000008EB00000A9CF0000C2F00000D1FF1100D8FF3100DEFF5100E3FF7100E9FF9100EFFFB100F6FFD100FFFFFF00000000002F26000050410000705B000090740000B08E0000CFA90000F0C30000FFD21100FFD83100FFDD5100FFE47100FFEA9100FFF0B100FFF6D100FFFFFF00000000002F1400005022000070300000903E0000B04D0000CF5B0000F0690000FF791100FF8A3100FF9D5100FFAF7100FFC19100FFD2B100FFE5D100FFFFFF00000000002F030000500400007006000090090000B00A0000CF0C0000F00E0000FF201200FF3E3100FF5C5100FF7A7100FF979100FFB6B100FFD4D100FFFFFF00000000002F000E00500017007000210090002B00B0003600CF004000F0004900FF115A00FF317000FF518600FF719C00FF91B200FFB1C800FFD1DF00FFFFFF00000000002F0020005000360070004C0090006200B0007800CF008E00F000A400FF11B300FF31BE00FF51C700FF71D100FF91DC00FFB1E500FFD1F000FFFFFF00000000002C002F004B0050006900700087009000A500B000C400CF00E100F000F011FF00F231FF00F451FF00F671FF00F791FF00F9B1FF00FBD1FF00FFFFFF00000000001B002F002D0050003F007000520090006300B0007600CF008800F0009911FF00A631FF00B451FF00C271FF00CF91FF00DCB1FF00EBD1FF00FFFFFF000000000008002F000E005000150070001B0090002100B0002600CF002C00F0003E11FF005831FF007151FF008C71FF00A691FF00BFB1FF00DAD1FF00FFFFFF0000001800040400083B4A260000000000001C0F3944371F0D0C4A26000000000000154360615A3C1F004D400000000000001545646462571D022122000000000000001A6F6F635D160E214B4F332D50000000013A5A5846051B2800312F2D360000000A185B591907230000302D2B51000000091E64633806545F002B2B2B52000000110B5A5E15204C41554E0000000000001310030017003F254953000000000000132A12140000292448530056000000003E42275C0000003D47530035353400000000000000000000000000322C2B000000000000000000000000002B2B2B000000000000000000000000002E2B2B00000000000000000000000000000000C01F0000801F0000801F0000801F0000C0010000C0210000C0610000C0210000C00F0000C10F0000C30B0000C3880000FFF80000FFF80000FFF80000FFFF000000
	END


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

	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(901,''Word'',''Word 97-2003 Document (*.doc)''           ,''doc'',     0, 0,1)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(902,''Word'',''Word 2007-2010 Document (*.docx)''        ,''docx'', null,16,0)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(903,''Word'',''XML document format (*.xml)''             ,''xml'',  null,12,0)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(904,''Word'',''PDF format (*.pdf)''                      ,''pdf'',  null,17,0)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(905,''Word'',''XPS format (*.xps)''                      ,''xps'',  null,18,0)'

	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(911,''WordTemplate'',''Word 97-2003 Document (*.doc)''   ,''doc'',     0, 0,0)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(912,''WordTemplate'',''Word 97-2003 Template (*.dot)''   ,''dot'',     0, 0,1)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(913,''WordTemplate'',''Word 2007-2010 Document (*.docx)'',''docx'', null, 0,0)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(914,''WordTemplate'',''Word 2007-2010 Template (*.dotx)'',''dotx'', null, 0,0)'

	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(921,''Excel'',''Excel 97-2003 Workbook (*.xls)''         ,''xls'', -4143,56,1)'
	EXEC sp_executesql N'INSERT ASRSysFileFormats VALUES(922,''Excel'',''Excel 2007-2010 Workbook (*.xlsx)''      ,''xlsx'', null,51,0)'


	IF NOT EXISTS(SELECT id FROM syscolumns WHERE id = OBJECT_ID('ASRSysCalendarReports', 'U') AND name = 'OutputSaveFormat')
	BEGIN
		EXEC sp_executesql
			N'ALTER TABLE ASRSysCalendarReports
            ADD OutputSaveFormat int NULL, OutputEmailFileFormat int NULL'
		EXEC sp_executesql
			N'UPDATE ASRSysCalendarReports
            SET OutputSaveFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END,
            OutputEmailFileFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE id = OBJECT_ID('ASRSysCrossTab', 'U') AND name = 'OutputSaveFormat')
	BEGIN
		EXEC sp_executesql
			N'ALTER TABLE ASRSysCrossTab
            ADD OutputSaveFormat int NULL, OutputEmailFileFormat int NULL'
		EXEC sp_executesql
			N'UPDATE ASRSysCrossTab
            SET OutputSaveFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END,
            OutputEmailFileFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE id = OBJECT_ID('ASRSysCustomReportsName', 'U') AND name = 'OutputSaveFormat')
	BEGIN
		EXEC sp_executesql N'ALTER TABLE ASRSysCustomReportsName
			ADD OutputSaveFormat int NULL, OutputEmailFileFormat int NULL'
		EXEC sp_executesql
			N'UPDATE ASRSysCustomReportsName
            SET OutputSaveFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END,
            OutputEmailFileFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE id = OBJECT_ID('ASRSysExportName', 'U') AND name = 'OutputSaveFormat')
	BEGIN
		EXEC sp_executesql
			N'ALTER TABLE ASRSysExportName
            ADD OutputSaveFormat int NULL, OutputEmailFileFormat int NULL'
		EXEC sp_executesql
			N'UPDATE ASRSysExportName
            SET OutputSaveFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END,
            OutputEmailFileFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE id = OBJECT_ID('ASRSysMatchReportName', 'U') AND name = 'OutputSaveFormat')
	BEGIN
		EXEC sp_executesql
			N'ALTER TABLE ASRSysMatchReportName
			ADD OutputSaveFormat int NULL, OutputEmailFileFormat int NULL'
		EXEC sp_executesql
			N'UPDATE ASRSysMatchReportName
            SET OutputSaveFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END,
            OutputEmailFileFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END'
	END

	IF NOT EXISTS(SELECT id FROM syscolumns WHERE id = OBJECT_ID('ASRSysRecordProfileName', 'U') AND name = 'OutputSaveFormat')
	BEGIN
		EXEC sp_executesql
			N'ALTER TABLE ASRSysRecordProfileName
			ADD OutputSaveFormat int NULL, OutputEmailFileFormat int NULL'
		EXEC sp_executesql
			N'UPDATE ASRSysRecordProfileName
            SET OutputSaveFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END,
            OutputEmailFileFormat = CASE WHEN OutputFormat = 3 THEN 921 WHEN OutputFormat IN (4,5,6) THEN 901 ELSE 0 END'
	END



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
