CREATE PROCEDURE dbo.spASRGetWorkflowGridItems
			(
				@piInstanceID		integer,
				@piElementItemID	integer, 
				@pfOK				bit	OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iTableID 		integer,
					@iOrderID			integer,
					@iFilterID			integer,
					@sFilterSQL			varchar(8000),
					@sFilterUDF			varchar(8000),
					@sRecSelWebFormIdentifier	varchar(200),
					@sRecSelIdentifier	varchar(200),
					@iDBRecord		integer,
					@iInitiatorID		integer,
					@sSQL			varchar(8000),
					@sOrderItemType	varchar(8000),
					@sSelectSQL		varchar(8000),
					@sOrderSQL		varchar(8000),
					@sBaseTableName	sysname,
					@fAscending		bit,
					@sColumnName		sysname,
					@sTempTableName	sysname,
					@iTempTableID		integer,
					@iTempTableType	integer,
					@iTempCount		integer,
					@sValue			varchar(8000),
					@iDataType		integer,
					@iRecordID		integer,
					@iPersonnelTableID	integer,
					@iWorkflowID		integer,
					@iElementType		integer, 
					@fValidRecordID	bit,
					@iElementID	integer,
					@iBaseTableID	integer,
					@iParent1TableID	int,
					@iParent1RecordID	int,
					@iParent2TableID	int,
					@iParent2RecordID	int,
					@iRecordTableID		int,
					@iTriggerTableID	integer
				DECLARE @joinParents table(tableID		integer)	
			
				SET @pfOK = 1

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
					@iTableID = ASRSysWorkflowElementItems.tableID,
					@iElementID = ASRSysWorkflowElementItems.elementiD,
					@sRecSelWebFormIdentifier = isnull(ASRSysWorkflowElementItems.wfFormIdentifier, ''),
					@sRecSelIdentifier = isnull(ASRSysWorkflowElementItems.wfValueIdentifier, 0),
					@iDBRecord = ASRSysWorkflowElementItems.dbRecord,
					@iOrderID = 
						CASE
							WHEN isnull(ASRSysWorkflowElementItems.recordOrderID, 0) > 0 THEN ASRSysWorkflowElementItems.recordOrderID
							ELSE ASRSysTables.defaultOrderID
						END,
					@iFilterID = isnull(ASRSysWorkflowElementItems.recordFilterID, 0),
					@iRecordTableID = ASRSysWorkflowElementItems.recordTableID,
					@sBaseTableName = ASRSysTables.tableName
				FROM ASRSysWorkflowElementItems
				INNER JOIN ASRSysTables ON ASRSysWorkflowElementItems.tableID = ASRSysTables.tableID
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID
			
				SELECT @iInitiatorID = ASRSysWorkflowInstances.initiatorID,
					@iWorkflowID = ASRSysWorkflowInstances.workflowID, 
					@iTriggerTableID = ASRSysWorkflows.baseTable,
					@iParent1TableID = ASRSysWorkflowInstances.parent1TableID,
					@iParent1RecordID = ASRSysWorkflowInstances.parent1RecordID,
					@iParent2TableID = ASRSysWorkflowInstances.parent2TableID,
					@iParent2RecordID = ASRSysWorkflowInstances.parent2RecordID
				FROM ASRSysWorkflowInstances
				INNER JOIN ASRSysWorkflows ON ASRSysWorkflowInstances.workflowID = ASRSysWorkflows.ID
				WHERE ASRSysWorkflowInstances.ID = @piInstanceID
			
				SET @sSelectSQL = ''
				SET @sOrderSQL = ''
			
				DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT 
					ASRSysColumns.columnName,
					ASRSysColumns.dataType,
					ASRSysColumns.tableID,
					ASRSysTables.tableType,
					ASRSysTables.tableName,
					upper(isnull(ASRSysOrderItems.type, '')),
					ASRSysOrderItems.ascending
				FROM ASRSysOrderItems
				INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
				INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
				WHERE ASRSysOrderItems.orderID = @iOrderID
				ORDER BY ASRSysOrderItems.type,
					ASRSysOrderItems.sequence
			
				OPEN orderCursor
				FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTempTableID, @iTempTableType, @sTempTableName, @sOrderItemType, @fAscending
				WHILE (@@fetch_status = 0)
				BEGIN
					IF @sOrderItemType = 'F'
					BEGIN
						SET @sSelectSQL = @sSelectSQL +
							CASE 
								WHEN len(@sSelectSQL) > 0 THEN ','
								ELSE ''
							END +
							@sTempTableName + '.' + @sColumnName
					END
			
					IF @sOrderItemType = 'O'
					BEGIN
						SET @sOrderSQL = @sOrderSQL + 
							CASE 
								WHEN len(@sOrderSQL) > 0 THEN ','
								ELSE ' '
							END + 
							@sTempTableName + '.' + @sColumnName +
							CASE 
								WHEN @fAscending = 0 THEN ' DESC' 
								ELSE '' 
							END				
					END
			
					IF @iTableID <> @iTempTableID
					BEGIN
						SELECT @iTempCount = COUNT(tableID)
						FROM @joinParents
						WHERE tableID = @iTempTableID
			
						IF @iTempCount = 0
						BEGIN
							INSERT INTO @joinParents (tableID) VALUES(@iTempTableID)
						END
					END
			
					FETCH NEXT FROM orderCursor INTO @sColumnName, @iDataType, @iTempTableID, @iTempTableType, @sTempTableName, @sOrderItemType, @fAscending
				END
				CLOSE orderCursor
				DEALLOCATE orderCursor
			
				IF len(@sSelectSQL) > 0 
				BEGIN
					SET @sSelectSQL = 'SELECT ' + @sSelectSQL + ',' +
						@sBaseTableName + '.id' +
					' FROM ' + @sBaseTableName
			
					DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT ASRSysTables.tableName, 
						JP.tableID
					FROM @joinParents JP
					INNER JOIN ASRSysTables ON JP.tableID = ASRSysTables.tableID
			
					OPEN joinCursor
					FETCH NEXT FROM joinCursor INTO @sTempTableName, @iTempTableID
					WHILE (@@fetch_status = 0)
					BEGIN
						SET @sSelectSQL = @sSelectSQL + 
							' LEFT OUTER JOIN ' + @sTempTableName + ' ON ' + @sBaseTableName + '.ID_' + convert(varchar(100), @iTempTableID) + ' = ' + @sTempTableName + '.ID'
			
						FETCH NEXT FROM joinCursor INTO @sTempTableName, @iTempTableID
					END
					CLOSE joinCursor
					DEALLOCATE joinCursor
			
					IF @iDBRecord = 0 -- ie. based on the initiator's record
					BEGIN
						SET @iBaseTableID = @iPersonnelTableID
						SET @iRecordID = @iInitiatorID
					END

					IF @iDBRecord = 4 -- ie. based on the triggered record
					BEGIN
						SET @iBaseTableID = @iTriggerTableID
						SET @iRecordID = @iInitiatorID
					END
			
					IF @iDBRecord = 1 -- ie. based on a previously identified record
					BEGIN
						SELECT @iElementType = ASRSysWorkflowElements.type
						FROM ASRSysWorkflowElements
						WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
							AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sRecSelWebFormIdentifier)))
		
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
								AND IV.elementID = Es.ID
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
							WHERE IV.instanceID = @piInstanceID
						END
			
						SET @iRecordID = 
							CASE
								WHEN isnumeric(@sValue) = 1 THEN convert(integer, @sValue)
								ELSE 0
							END

						SET @iBaseTableID = @iTempTableID
					END
			
					IF (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4)
					BEGIN
						EXEC [dbo].[spASRWorkflowAscendantRecordID]
							@iBaseTableID,
							@iRecordID,
							@iParent1TableID,
							@iParent1RecordID,
							@iParent2TableID,
							@iParent2RecordID,
							@iRecordTableID,
							@iRecordID	OUTPUT

						SET @sSelectSQL = @sSelectSQL + 
							' WHERE ' + @sBaseTableName + '.ID_' + convert(varchar(100), @iRecordTableID) + ' = ' + convert(varchar(100), @iRecordID)

						SET @fValidRecordID = 1

						EXEC [dbo].[spASRWorkflowValidTableRecord]
							@iRecordTableID,
							@iRecordID,
							@fValidRecordID	OUTPUT

						IF @fValidRecordID  = 0
						BEGIN
							SET @pfOK = 0

							-- Update the ASRSysWorkflowInstanceSteps table to show that this step has failed. 
							EXEC [dbo].[spASRWorkflowActionFailed] @piInstanceID, @iElementID, 'Web Form record selector item record has been deleted or not selected.'
							
							-- Need to return a recordset of some kind.
							SELECT '' AS 'Error'

							RETURN
						END
					END

					IF @iFilterID > 0 
					BEGIN
						SET @sFilterUDF = '[dbo].udf_ASRWFExpr_' + convert(varchar(8000), @iFilterID)

						IF EXISTS(
							SELECT Name
							FROM sysobjects
							WHERE id = object_id(@sFilterUDF)
							AND sysstat & 0xf = 0)
						BEGIN
							SET @sFilterSQL = 
								CASE
									WHEN (@iDBRecord = 0) OR (@iDBRecord = 1) OR (@iDBRecord = 4) THEN ' AND '
									ELSE ' WHERE '
								END 
								+ @sBaseTableName + '.ID  IN (SELECT id FROM ' + @sFilterUDF + '(' + convert(varchar(8000), @piInstanceID) + '))'
						END
					END

					SET @sOrderSQL = ' ORDER BY ' + @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sBaseTableName + '.ID'

					EXEC (@sSelectSQL 
						+ @sFilterSQL
						+ @sOrderSQL)
				END
			END
