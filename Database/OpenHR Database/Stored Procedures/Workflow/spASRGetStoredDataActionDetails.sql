CREATE PROCEDURE [dbo].[spASRGetStoredDataActionDetails]
(
	@piInstanceID		integer,
	@piElementID		integer,
	@psSQL				varchar(MAX)	OUTPUT, 
	@piDataTableID		integer			OUTPUT,
	@psTableName		varchar(255)	OUTPUT,
	@piDataAction		integer			OUTPUT, 
	@piRecordID			integer			OUTPUT,
	@bUseAsTargetIdentifier	bit OUTPUT,
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
      @maxSize                float,
		@iSize						integer,
		@iDecimals					integer,
		@iTriggerTableID			integer;
			
	SET @psSQL = '';
	SET @pfResult = 1;
	SET @piRecordID = 0;

	SELECT @iPersonnelTableID = convert(integer, ISNULL(parameterValue, '0'))
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_PERSONNEL'
		AND parameterKey = 'Param_TablePersonnel';

	IF @iPersonnelTableID = 0
	BEGIN
		SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, 0))
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_WORKFLOW'
		AND parameterKey = 'Param_TablePersonnel';
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
		@iTriggerTableID = ASRSysWorkflows.baseTable,
		@bUseAsTargetIdentifier = ISNULL(UseAsTargetIdentifier, 0)
	FROM ASRSysWorkflowElements
	INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID
	WHERE ASRSysWorkflowElements.ID = @piElementID;

	SELECT @psTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @piDataTableID;

	IF @iDataRecord = 0 -- 0 = Initiator's record
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
			SET @sIDColumnName = 'ID';
		END
		ELSE
		BEGIN
			SET @sIDColumnName = 'ID_' + convert(varchar(255), @iDataRecordTableID);
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
			SET @sIDColumnName = 'ID';
		END
		ELSE
		BEGIN
			SET @sIDColumnName = 'ID_' + convert(varchar(255), @iDataRecordTableID);
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
			SELECT @sValue = ISNULL(IV.value, '0'),
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
			SELECT @sValue = ISNULL(IV.value, '0'),
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
			SET @sIDColumnName = 'ID';
		END
		ELSE
		BEGIN
			SET @sIDColumnName = 'ID_' + convert(varchar(255), @iDataRecordTableID);
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
				'Stored Data primary record has been deleted or not selected.';

			SET @psSQL = '';
			SET @pfResult = 0;
			RETURN;
		END
	END

	IF @piDataAction = 0 -- Insert
	BEGIN
		IF @iSecondaryDataRecord = 0 -- 0 = Initiator's record
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
				SET @sSecondaryIDColumnName = 'ID';
			END
			ELSE
			BEGIN
				SET @sSecondaryIDColumnName = 'ID_' + convert(varchar(255), @iSecondaryDataRecordTableID);
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
				SET @sSecondaryIDColumnName = 'ID';
			END
			ELSE
			BEGIN
				SET @sSecondaryIDColumnName = 'ID_' + convert(varchar(255), @iSecondaryDataRecordTableID);
			END
		END

		IF @iSecondaryDataRecord = 1 -- 1 = Previous record selector's record
		BEGIN
			SELECT @iElementType = ASRSysWorkflowElements.type
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
				AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sSecondaryRecSelWebFormIdentifier)));
	
			IF @iElementType = 2
			BEGIN
				 -- WebForm
				SELECT @sValue = ISNULL(IV.value, '0'),
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
				SELECT @sValue = ISNULL(IV.value, '0'),
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
				SET @sSecondaryIDColumnName = 'ID';
			END
			ELSE
			BEGIN
				SET @sSecondaryIDColumnName = 'ID_' + convert(varchar(255), @iSecondaryDataRecordTableID);
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
					'Stored Data secondary record has been deleted or not selected.';

				SET @psSQL = '';
				SET @pfResult = 0;
				RETURN;
			END
		END

	END

	IF @piDataAction = 0 OR @piDataAction = 1
	BEGIN
		/* INSERT or UPDATE. */
		SET @sColumnList = '';
		SET @sValueList = '';

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
			''
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

			SET @sSQL = 'SELECT @sDBValue = '
				+ CASE
					WHEN @iDataType = 12 THEN ''
					WHEN @iDataType = 11 THEN 'convert(varchar(MAX),'
					ELSE 'convert(varchar(MAX),'
				END
				+ @sDBTableName + '.' + @sDBColumnName
				+ CASE
					WHEN @iDataType = 12 THEN ''
					WHEN @iDataType = 11 THEN ', 101)'
					ELSE ')'
				END
				+ ' FROM ' + @sDBTableName 
				+ ' WHERE ' + @sDBTableName + '.ID = ';

			SET @iRecordID = 0;

			IF @iDBRecord = 0
			BEGIN
				-- Initiator's record
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
						AND IV.elementID = Es.ID;
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
							SELECT @sDBValue = rtrim(ltrim(isnull(QC.columnValue , '')))
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
								SELECT @sDBValue = rtrim(ltrim(isnull(IV.value , '')))
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
						'Stored Data column database value record has been deleted or not selected.';

					SET @psSQL = '';
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
					SET @sParam = N'@sDBValue varchar(MAX) OUTPUT';
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
								WHEN UPPER(EC.value) = 'TRUE' THEN '1'
								ELSE '0'
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
				ELSE '' -- Database Value. Handle below to avoid collation conflict.
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
					SET @maxSize = convert(float, '1' + REPLICATE('0', @iSize - @iDecimals))
					IF @fltResult >= @maxSize SET @fltResult = 0;
					IF @fltResult <= (-1 * @maxSize) SET @fltResult = 0;
				END

				SET @sValue = 
					CASE
						WHEN @iResultType = 2 THEN ltrim(rtrim(STR(@fltResult, 8000, @iDecimals)))
						WHEN @iResultType = 3 THEN 
							CASE 
								WHEN @fResult = 1 THEN '1'
								ELSE '0'
							END
						WHEN (@iResultType = 4) THEN
							CASE 
								WHEN @dtResult is NULL THEN 'NULL'
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
						WHEN LEN(@sColumnList) > 0 THEN ','
						ELSE ''
					END
					+ @sColumnName;

				SET @sValueList = @sValueList
					+ CASE
						WHEN LEN(@sValueList) > 0 THEN ','
						ELSE ''
					END
					+ CASE
						WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''' + replace(isnull(@sValue, ''), '''', '''''') + '''' -- 12 = varchar, -1 = working pattern
						WHEN @iColumnDataType = 11 THEN
							CASE 
								WHEN (upper(ltrim(rtrim(@sValue))) = 'NULL') OR (@sValue IS null) THEN 'null'
								ELSE '''' + replace(@sValue, '''', '''''') + '''' -- 11 = date
							END
						WHEN LEN(@sValue) = 0 THEN '0'
						ELSE isnull(@sValue, 0) -- integer, logic, numeric
					END;
			END
			ELSE
			BEGIN
				/* UPDATE. */
				SET @sColumnList = @sColumnList
					+ CASE
						WHEN LEN(@sColumnList) > 0 THEN ','
						ELSE ''
					END
					+ @sColumnName
					+ ' = '
					+ CASE
						WHEN @iColumnDataType = 12 OR @iColumnDataType = -1 THEN '''' + replace(isnull(@sValue, ''), '''', '''''') + '''' -- 12 = varchar, -1 = working pattern
						WHEN @iColumnDataType = 11 THEN
							CASE 
								WHEN (upper(ltrim(rtrim(@sValue))) = 'NULL') OR (@sValue IS null) THEN 'null'
								ELSE '''' + replace(@sValue, '''', '''''') + '''' -- 11 = date
							END
						WHEN LEN(@sValue) = 0 THEN '0'
						ELSE isnull(@sValue, 0) -- integer, logic, numeric
					END;
			END

			DELETE FROM [dbo].[ASRSysWorkflowInstanceValues]
			WHERE instanceID = @piInstanceID
				AND elementID = @piElementID
				AND columnID = @iColumnID;

			INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
				(instanceID, elementID, identifier, columnID, value, emailID)
				VALUES (@piInstanceID, @piElementID, '', @iColumnID, @sValue, 0);

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
						WHEN LEN(@sColumnList) > 0 THEN ','
						ELSE ''
					END
					+ @sIDColumnName;
	
				SET @sValueList = @sValueList
					+ CASE
						WHEN LEN(@sValueList) > 0 THEN ','
						ELSE ''
					END
					+ convert(varchar(255), @piRecordID);

				IF @piDataAction = 0 -- Insert
					AND (@iSecondaryDataRecord = 0 -- 0 = Initiator's record
						OR @iSecondaryDataRecord = 1 -- 1 = Previous record selector's record
						OR @iSecondaryDataRecord = 4) -- 4 = Triggered record
				BEGIN
					SET @sColumnList = @sColumnList
						+ CASE
							WHEN LEN(@sColumnList) > 0 THEN ','
							ELSE ''
						END
						+ @sSecondaryIDColumnName;
				
					SET @sValueList = @sValueList
						+ CASE
							WHEN LEN(@sValueList) > 0 THEN ','
							ELSE ''
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
				SET @psSQL = 'INSERT INTO ' + @psTableName
					+ ' (' + @sColumnList + ')'
					+ ' VALUES(' + @sValueList + ')';
				SET @pfResult = 1;
			END
			ELSE
			BEGIN
				/* UPDATE. */
				SET @psSQL = 'UPDATE ' + @psTableName
					+ ' SET ' + @sColumnList
					+ ' WHERE ' + @sIDColumnName + ' = ' + convert(varchar(255), @piRecordID);
				SET @pfResult = 1;
			END
		END
	END

	IF @piDataAction = 2
	BEGIN
		/* DELETE. */
		SET @psSQL = 'DELETE FROM ' + @psTableName
			+ ' WHERE ' + @sIDColumnName + ' = ' + convert(varchar(255), @piRecordID);
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

			SET @sSQL = 'SELECT @sDBValue = '
				+ CASE
					WHEN @iDataType = 12 THEN ''
					WHEN @iDataType = 11 THEN 'convert(varchar(MAX),'
					ELSE 'convert(varchar(MAX),'
				END
				+ @sDBTableName + '.' + @sDBColumnName
				+ CASE
					WHEN @iDataType = 12 THEN ''
					WHEN @iDataType = 11 THEN ', 101)'
					ELSE ')'
				END
				+ ' FROM ' + @sDBTableName 
				+ ' WHERE ' + @sDBTableName + '.ID = ' + convert(varchar(255), @piRecordID);

			SET @sParam = N'@sDBValue varchar(MAX) OUTPUT';
			EXEC sp_executesql @sSQL, @sParam, @sDBValue OUTPUT;

			INSERT INTO [dbo].[ASRSysWorkflowInstanceValues]
				(instanceID, elementID, identifier, columnID, value, emailID)
				VALUES (@piInstanceID, @piElementID, '', @iDBColumnID, @sDBValue, 0);
					
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

				SET @sSQL = 'SELECT @sDBValue = '
					+ CASE
						WHEN @iDataType = 12 THEN ''
						WHEN @iDataType = 11 THEN 'convert(varchar(MAX),'
						ELSE 'convert(varchar(MAX),'
					END
					+ @sDBTableName + '.' + @sDBColumnName
					+ CASE
						WHEN @iDataType = 12 THEN ''
						WHEN @iDataType = 11 THEN ', 101)'
						ELSE ')'
					END
					+ ' FROM ' + @sDBTableName 
					+ ' WHERE ' + @sDBTableName + '.ID = ' + convert(varchar(255), @piRecordID);

				SET @sParam = N'@sDBValue varchar(MAX) OUTPUT';
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
				VALUES (@piInstanceID, @piElementID, '', 0, @sDBValue, @iEmailID);
					
			FETCH NEXT FROM curEmails INTO @iEmailID, @iType, @iDBColumnID;
		END
		CLOSE curEmails;
		DEALLOCATE curEmails;
	END
END