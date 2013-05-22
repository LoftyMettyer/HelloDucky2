CREATE PROCEDURE [dbo].[spASRGetWorkflowFormItems]
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
		SET @psErrorMessage = 'This workflow step is invalid. The workflow process may have been completed.'
		RETURN
	END

	-- Check if the step has already been completed!
	SELECT @iStatus = ASRSysWorkflowInstanceSteps.status
	FROM ASRSysWorkflowInstanceSteps
	WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
		AND ASRSysWorkflowInstanceSteps.elementID = @piElementID

	IF @iStatus = 3
	BEGIN
		SET @psErrorMessage = 'This workflow step has already been completed.'
		RETURN
	END

	IF @iStatus = 6
	BEGIN
		SET @psErrorMessage = 'This workflow step has timed out.'
		RETURN
	END

	IF @iStatus = 0
	BEGIN
		SET @psErrorMessage = 'This workflow step is invalid. It may no longer be required due to the results of other workflow steps.'
		RETURN
	END

	SET @psErrorMessage = ''

	SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, '0'))
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_PERSONNEL'
		AND parameterKey = 'Param_TablePersonnel'

	IF @iPersonnelTableID = 0
	BEGIN
		SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_WORKFLOW'
		AND parameterKey = 'Param_TablePersonnel'
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
		SET @sValue = ''

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
				-- Initiator's record
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
					SELECT @sValue = ISNULL(IV.value, '0'),
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
					SELECT @sValue = ISNULL(IV.value, '0'),
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
							SELECT @sValue = rtrim(ltrim(isnull(QC.columnValue , '')))
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
								SELECT @sValue = rtrim(ltrim(isnull(IV.value , '')))
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
					EXEC [dbo].[spASRWorkflowActionFailed] @piInstanceID, @piElementID, 'Web Form item record has been deleted or not selected.'
								
					SET @psErrorMessage = 'Error loading web form. Web Form item record has been deleted or not selected.'
					RETURN
				END
			END
				
			IF @fDeletedValue = 0
			BEGIN
				IF @iDBColumnDataType = 11 -- Date column, need to format into MM\DD\YYYY
				BEGIN
					SET @sSQL = 'SELECT @sValue = convert(varchar(100), ' + @sTableName + '.' + @sColumnName + ', 101)'
				END
				ELSE
				BEGIN
					SET @sSQL = 'SELECT @sValue = ' + @sTableName + '.' + @sColumnName
				END
				
				SET @sSQL = @sSQL +
						' FROM ' + @sTableName +
						' WHERE ' + @sTableName + '.ID = ' + convert(nvarchar(100), @iRecordID)
				SET @sSQLParam = N'@sValue varchar(MAX) OUTPUT'
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
						WHEN (@iItemType = 6 AND IVs.value = '1') THEN 'TRUE' 
						WHEN (@iItemType = 6 AND IVs.value <> '1') THEN 'FALSE' 
						WHEN (@iItemType = 7 AND (upper(ltrim(rtrim(IVs.value))) = 'NULL')) THEN '' 
						WHEN (@iItemType = 17 AND IVs.fileUpload_File IS null) THEN '0'
						WHEN (@iItemType = 17 AND NOT IVs.fileUpload_File IS null) THEN '1'
						ELSE isnull(IVs.value, '')
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
									WHEN @fResult = 1 THEN 'TRUE'
									ELSE 'FALSE'
								END
							WHEN @iResultType = 4 THEN convert(varchar(100), @dtResult, 101)
							ELSE convert(varchar(MAX), @sResult)
						END

					SET @iType = @iResultType
				END
				ELSE
				BEGIN
					SELECT @sValue = isnull(EIs.inputDefault, '')
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
		IV.type AS [sourceItemType],
		LUFC.ColumnName AS [lookupFilterColumnName],
		LUFC.datatype AS [lookupFilterColumnDataType],
		LUI.ID AS [lookupFilterValueID],
		LUI.ItemType AS [lookupFilterValueType]
	FROM ASRSysWorkflowElementItems thisFormItems
	LEFT OUTER JOIN @itemValues IV ON thisFormItems.ID = IV.ID
	LEFT OUTER JOIN ASRSysColumns LUFC ON thisFormItems.lookupFilterColumnID = LUFC.ColumnID
	LEFT OUTER JOIN ASRSysWorkflowElementItems LUI ON thisFormItems.lookupFilterValue = LUI.Identifier
		AND LUI.elementID = @piElementID
		AND LEN(LUI.Identifier) > 0
	WHERE thisFormItems.elementID = @piElementID
	ORDER BY thisFormItems.ZOrder DESC
END
