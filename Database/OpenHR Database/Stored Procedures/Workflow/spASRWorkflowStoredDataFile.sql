CREATE PROCEDURE [dbo].[spASRWorkflowStoredDataFile]
(
	@piElementColumnID	integer,
	@piInstanceID		integer,
	@piValueType		integer			OUTPUT,
	@psFileName			varchar(MAX)	OUTPUT,
	@psErrorMessage		varchar(MAX)	OUTPUT,
	@piOLEType			integer			OUTPUT
)
AS
BEGIN
	DECLARE 
		@iWorkflowID			integer,
		@iElementID				integer,
		@sElementIdentifier		varchar(MAX),
		@sItemIdentifier		varchar(MAX),
		@iDBColumnID			integer,
		@iDBRecord				integer,
		@sTableName				sysname,
		@sColumnName			sysname,
		@iRequiredTableID		integer,
		@iRequiredRecordID		integer,
		@iRecordID				integer,
		@iBaseTableID			integer,
		@iBaseRecordID			integer,
		@iParent1TableID		integer,
		@iParent1RecordID		integer,
		@iParent2TableID		integer,
		@iParent2RecordID		integer,
		@iInitiatorID			integer,
		@iInitParent1TableID	integer,
		@iInitParent1RecordID	integer,
		@iInitParent2TableID	integer,
		@iInitParent2RecordID	integer,
		@iElementType			integer, 
		@iTempElementID			integer,
		@sValue					varchar(MAX),
		@fValidRecordID			bit,
		@fDeletedValue			bit,
		@iPersonnelTableID		integer,
		@iCount					integer,
		@sSQL					nvarchar(MAX),
		@sSQLParam				nvarchar(MAX);

	SELECT @iWorkflowID = isnull(WE.workflowID, 0),
		@iBaseTableID = isnull(WF.baseTable, 0),
		@piValueType = isnull(WEC.valueType, 0),
		@sElementIdentifier = upper(rtrim(ltrim(isnull(WEC.WFFormIdentifier, '')))),
		@sItemIdentifier = upper(rtrim(ltrim(isnull(WEC.WFValueIdentifier, '')))),
		@iDBColumnID = isnull(WEC.DBColumnID, 0),
		@iDBRecord = isnull(WEC.DBRecord, 0)
	FROM ASRSysWorkflowElementColumns WEC
	INNER JOIN ASRSysWorkflowElements WE ON WEC.elementID = WE.ID
	INNER JOIN ASRSysWorkflows WF ON WE.workflowID = WF.ID
	WHERE WEC.ID = @piElementColumnID;

	IF @piValueType = 2 -- DB File
	BEGIN
		SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
			@iInitParent1TableID = ASRSysWorkflowInstances.parent1TableID,
			@iInitParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
			@iInitParent2TableID = ASRSysWorkflowInstances.parent2TableID,
			@iInitParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
		FROM ASRSysWorkflowInstances
		WHERE ASRSysWorkflowInstances.ID = @piInstanceID;

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

		SET @fDeletedValue = 0;

		SELECT @sTableName = ASRSysTables.tableName, 
			@iRequiredTableID = ASRSysTables.tableID, 
			@sColumnName = ASRSysColumns.columnName,
			@piOLEType = ASRSysColumns.OLEType
		FROM ASRSysColumns
		INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
		WHERE ASRSysColumns.columnID = @iDBColumnID;

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
		END

		IF @iDBRecord = 1
		BEGIN
			-- Identified record.
			SELECT @iElementType = ASRSysWorkflowElements.type, 
				@iTempElementID = ASRSysWorkflowElements.ID
			FROM ASRSysWorkflowElements
			WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
				AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sElementIdentifier)));
				
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
					AND IV.identifier = @sItemIdentifier
					AND Es.identifier = @sElementIdentifier
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
					AND Es.identifier = @sElementIdentifier
				WHERE IV.instanceID = @piInstanceID;
			END

			SET @iRecordID = 
				CASE
					WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
					ELSE 0
				END
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
				IF @iDBRecord = 4 -- Trigger record. 
				BEGIN
					SELECT @iCount = COUNT(*)
					FROM ASRSysWorkflowQueueColumns QC
					INNER JOIN ASRSysWorkflowQueue WFQ ON QC.queueID = WFQ.queueID
					WHERE WFQ.instanceID = @piInstanceID
						AND QC.columnID = @iDBColumnID;

					IF @iCount = 1
					BEGIN
						SELECT @sValue = rtrim(ltrim(isnull(QC.columnValue , '')))
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
							SELECT @sValue = rtrim(ltrim(isnull(IV.value , '')))
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
				SET @psErrorMessage = 'Record has been deleted or not selected.';
				RETURN;
			END
		END
			
		IF @fDeletedValue = 0
		BEGIN
			IF (@piOLEType = 0) OR (@piOLEType = 1)
			BEGIN
				SET @sSQL = 'SELECT @psFileName = ' + @sTableName + '.' + @sColumnName +
					' FROM ' + @sTableName +
					' WHERE ' + @sTableName + '.ID = ' + convert(nvarchar(255), @iRecordID);
				SET @sSQLParam = N'@psFileName varchar(MAX) OUTPUT';
				EXEC sp_executesql @sSQL, @sSQLParam, @psFileName OUTPUT;
			END
			ELSE
			BEGIN
				SET @sSQL = 'SELECT ' + @sTableName + '.' + @sColumnName + ' AS [file]' +
					' FROM ' + @sTableName +
					' WHERE ' + @sTableName + '.ID = ' + convert(nvarchar(255), @iRecordID);
				EXEC sp_executesql @sSQL;
			END
		END
	END
	
	IF @piValueType = 1 -- WF File
	BEGIN
		SELECT @iElementID = isnull(ID, 0)
		FROM ASRSysWorkflowElements
		WHERE workflowID = @iWorkflowID
			AND upper(ltrim(rtrim(isnull(identifier, '')))) = @sElementIdentifier;

		SELECT @psFileName = fileUpload_fileName
		FROM ASRSysWorkflowInstanceValues
		WHERE instanceID = @piInstanceID
			AND elementID = @iElementID
			AND upper(ltrim(rtrim(isnull(identifier, '')))) = @sItemIdentifier;

		SELECT fileUpload_file AS [file]
		FROM ASRSysWorkflowInstanceValues
		WHERE instanceID = @piInstanceID
			AND elementID = @iElementID
			AND upper(ltrim(rtrim(isnull(identifier, '')))) = @sItemIdentifier;
	END
END