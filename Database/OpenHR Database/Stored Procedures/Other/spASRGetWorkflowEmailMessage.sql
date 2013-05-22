CREATE PROCEDURE [dbo].[spASRGetWorkflowEmailMessage]
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
	SET @psMessage = '';
	SET @psMessage_HypertextLinks = '';
	SET @psHypertextLinkedSteps = '';
	SELECT @iSQLVersion = dbo.udfASRSQLVersion();

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

	exec [dbo].[spASRGetSetting]
		'email',
		'date format',
		'103',
		0,
		@sEmailFormat OUTPUT;

	SET @iEmailFormat = convert(integer, @sEmailFormat);
	
	SELECT @sURL = parameterValue
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_WORKFLOW'
		AND parameterKey = 'Param_URL';

	IF upper(right(@sURL, 5)) <> '.ASPX'
		AND right(@sURL, 1) <> '/'
		AND len(@sURL) > 0
	BEGIN
		SET @sURL = @sURL + '/';
	END

	SELECT @sParam1 = parameterValue
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_WORKFLOW'		
		AND parameterKey = 'Param_Web1';
	
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
		SET @sValue = '';

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
				-- Previously identified record.
				SELECT @iElementType = ASRSysWorkflowElements.type, 
					@iTempElementID = ASRSysWorkflowElements.ID
				FROM ASRSysWorkflowElements
				WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
					AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)));

				IF @iElementType = 2
				BEGIN
					 -- WebForm
					SELECT @sTemp = ISNULL(IV.value, '0'),
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
					SELECT @sTemp = ISNULL(IV.value, '0'),
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

				IF @fValidRecordID  = 0
				BEGIN
					SET @psMessage = '';
					SET @pfOK = 0;

					RETURN;
				END
			END

			IF @fDeletedValue = 0
			BEGIN
				SET @sSQL = 'SELECT @sValue = ' + @sTableName + '.' + @sColumnName +
					' FROM ' + @sTableName +
					' WHERE ' + @sTableName + '.ID = ' + convert(nvarchar(255), @iRecordID);
				SET @sSQLParam = N'@sValue varchar(MAX) OUTPUT';
				EXEC sp_executesql @sSQL, @sSQLParam, @sValue OUTPUT;
			END					
			IF @sValue IS null SET @sValue = '';

			/* Format dates */
			IF @iSourceItemType = 11
			BEGIN
				IF (len(@sValue) = 0) OR (@sValue = 'null')
				BEGIN
					SET @sValue = '<undefined>';
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
					SET @sValue = 'False';
				END
				ELSE
				BEGIN
					SET @sValue = 'True';
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

			IF @sValue IS null SET @sValue = '';

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
				IF len(@sValue) = 0 OR @sValue = 'null'
				BEGIN
					SET @sValue = '<undefined>';
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
					SET @sValue = 'False';
				END
				ELSE
				BEGIN
					SET @sValue = 'True';
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
					WHEN @sCaption = 'L' THEN '' + char(13) + char(10) + '--------------------------------------------------' + char(13) + char(10)
					WHEN @sCaption = 'N' THEN '' + char(13) + char(10)
					WHEN @sCaption = 'T' THEN '' + char(9)
					ELSE ''
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
			+ 'Click on the following link'
			+ CASE
				WHEN @iCount = 1 THEN ''
				ELSE 's'
			END
			+ ' to action:'
			+ char(13) + char(10);

		DECLARE elementCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT SE.elementID, ISNULL(WE.caption, '')
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
					@sCaption + ' - Error constructing the query string. Please contact your system administrator.';
			END
			ELSE
			BEGIN
				SET @psHypertextLinkedSteps = @psHypertextLinkedSteps
					+ CASE
						WHEN len(@psHypertextLinkedSteps) = 0 THEN char(9)
						ELSE ''
					END 
					+ convert(varchar(MAX), @iElementID)
					+ char(9);

				SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + char(13) + char(10) +
					@sCaption + ' - ' + char(13) + char(10) + 
					'<' + @sURL + '?' + @sQueryString + '>';
			END
			
			FETCH NEXT FROM elementCursor INTO @iElementID, @sCaption;
		END

		CLOSE elementCursor;
		DEALLOCATE elementCursor;

		SET @psMessage_HypertextLinks = @psMessage_HypertextLinks + char(13) + char(10) + char(13) + char(10)
			+ 'Please make sure that the link'
			+ CASE
				WHEN @iCount = 1 THEN ' has'
				ELSE 's have'
			END
			+ ' not been cut off by your display.' + char(13) + char(10)
			+ 'If '
			+ CASE
				WHEN @iCount = 1 THEN 'it has'
				ELSE 'they have'
			END
			+ ' been cut off you will need to copy and paste '
			+ CASE
				WHEN @iCount = 1 THEN 'it'
				ELSE 'them'
			END
			+ ' into your browser.';
	END

	DROP TABLE #succeedingElements;
END
