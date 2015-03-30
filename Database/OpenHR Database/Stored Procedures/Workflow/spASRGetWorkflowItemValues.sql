CREATE PROCEDURE [dbo].[spASRGetWorkflowItemValues]
			(
				@piElementItemID	integer,
				@piInstanceID	integer, 
				@piLookupColumnIndex	integer OUTPUT, 
				@piItemType	integer OUTPUT, 
				@psDefaultValue	varchar(8000) OUTPUT
			)
			AS
			BEGIN
				DECLARE 
					@iItemType			integer,
					@iResultType	integer,
					@sResult		varchar(8000),
					@fResult		bit,
					@dtResult		datetime,
					@fltResult		float,
					@iDefaultValueType		integer,
					@iCalcID				integer,
					@iLookupColumnID	integer,
					@sDefaultValue		varchar(8000),
					@sTableName			sysname,
					@sColumnName		sysname,
					@iDataType			integer,
					@iOrderID			integer,
					@iTableID			integer,
					@sSelectSQL			varchar(max),
					@sColumnList		varchar(max),
					@sOrderSQL			varchar(max),
					@sJoinSQL			varchar(max),
					@sJoinedTables		varchar(max),
					@fLookupColumnDoneF	bit,
					@sOrderType	char(1),
					@fOrderAsc	bit,
					@sOrderTableName	sysname,
					@sOrderColumnName	sysname,
					@iOrderColumnID	integer,
					@iOrderTableID	integer,
					@sTemp	varchar(max),
					@iCount	integer,
					@iStatus			integer,
					@iElementID			integer,
					@sValue				varchar(8000),
					@sIdentifier		varchar(8000),
					@sLookupFilterColumnName	varchar(8000),
					@iLookupFilterColumnType	int,
					@iLookupOrderID		int;

				SET @piLookupColumnIndex = 0;
								
				DECLARE @dropdownValues TABLE([value] varchar(255));

				SELECT 			
					@iItemType = ASRSysWorkflowElementItems.itemType,
					@sDefaultValue = ASRSysWorkflowElementItems.inputDefault,
					@iLookupColumnID = ASRSysWorkflowElementItems.lookupColumnID,
					@iElementID = ASRSysWorkflowElementItems.elementID,
					@sIdentifier = ASRSysWorkflowElementItems.identifier,
					@iCalcID = isnull(ASRSysWorkflowElementItems.calcID, 0),
					@iDefaultValueType = isnull(ASRSysWorkflowElementItems.defaultValueType, 0),
					@sLookupFilterColumnName = isnull(COLS.columnName, ''),
					@iLookupFilterColumnType = isnull(COLS.dataType, 0),
					@iLookupOrderID = ASRSysWorkflowElementItems.LookupOrderID
				FROM ASRSysWorkflowElementItems
				LEFT OUTER JOIN ASRSysColumns COLS ON ASRSysWorkflowElementItems.LookupFilterColumnID = COLS.columnID
				WHERE ASRSysWorkflowElementItems.ID = @piElementItemID;

				SET @piItemType = @iItemType;

				SELECT @iStatus = ASRSysWorkflowInstanceSteps.status
				FROM ASRSysWorkflowInstanceSteps
				WHERE ASRSysWorkflowInstanceSteps.instanceID = @piInstanceID
					AND ASRSysWorkflowInstanceSteps.elementID = @iElementID;

				IF @iStatus = 7 -- Previously SavedForLater
				BEGIN
					SELECT @sValue = isnull(IVs.value, '')
					FROM ASRSysWorkflowInstanceValues IVs
					WHERE IVs.instanceID = @piInstanceID
						AND IVs.elementID = @iElementID
						AND IVs.identifier = @sIdentifier;

					SET @sDefaultValue = @sValue;
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
							0;

						SET @sDefaultValue = 
							CASE
								WHEN @iResultType = 2 THEN convert(varchar(8000), @fltResult)
								WHEN @iResultType = 3 THEN 
									CASE 
										WHEN @fResult = 1 THEN 'TRUE'
										ELSE 'FALSE'
									END
								WHEN @iResultType = 4 THEN convert(varchar(100), @dtResult, 101)
								ELSE convert(varchar(8000), @sResult)
							END;
					END
				END

				SET @psDefaultValue = @sDefaultValue;

				IF @iItemType = 15 -- OptionGroup
				BEGIN
					SELECT ASRSysWorkflowElementItemValues.value,
						CASE
							WHEN ASRSysWorkflowElementItemValues.value = @sDefaultValue THEN 1
							ELSE 0
						END AS [ASRSysDefaultValueFlag]
					FROM ASRSysWorkflowElementItemValues
					WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
					ORDER BY ASRSysWorkflowElementItemValues.sequence;
				END

				IF @iItemType = 13 -- Dropdown
				BEGIN
					INSERT INTO @dropdownValues ([value])
						SELECT ASRSysWorkflowElementItemValues.value
						FROM ASRSysWorkflowElementItemValues
						WHERE ASRSysWorkflowElementItemValues.itemID = @piElementItemID
						ORDER BY [sequence];

					SELECT [value],
						'' AS [ASRSysLookupFilterValue]				
					FROM @dropdownValues;
				END
				
				IF (@iItemType = 14) AND (@iLookupColumnID > 0) -- Lookup
				BEGIN
					SELECT @sTableName = ASRSysTables.tableName,
						@sColumnName = ASRSysColumns.columnName,
						@iOrderID = COALESCE(NULLIF(@iLookupOrderID, 0), ASRSysTables.defaultOrderID),
						@iTableID = ASRSysTables.tableID,
						@iDataType = ASRSysColumns.dataType
					FROM ASRSysColumns
					INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID
					WHERE ASRSysColumns.columnId = @iLookupColumnID;		

					IF @iDataType = 11 -- Date 
						AND UPPER(LTRIM(RTRIM(@sDefaultValue))) = 'NULL'
					BEGIN
						SET @sDefaultValue = '';
					END

					SET @sColumnList = '';
					SET @sJoinSQL ='';
					SET @sOrderSQL = '';
					SET @fLookupColumnDoneF = 0;
					SET @sJoinedTables = ',';
					SET @iCount = 0;
				
					DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT ASRSysOrderItems.type,
						ASRSysTables.tableName,
						ASRSysColumns.columnName,
						ASRSysColumns.columnId,
						ASRSysColumns.tableID,
						ASRSysOrderItems.ascending
					FROM ASRSysOrderItems
					INNER JOIN ASRSysColumns 
						ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
					INNER JOIN ASRSysTables 
						ON ASRSysTables.tableID = ASRSysColumns.tableID
					WHERE ASRSysOrderItems.orderID = @iOrderID
					AND ASRSysColumns.datatype <> -3
					AND ASRSysColumns.datatype <> -4
					ORDER BY ASRSysOrderItems.type, 
						ASRSysOrderItems.sequence;

					OPEN orderCursor;
					FETCH NEXT FROM orderCursor INTO 
						@sOrderType, 
						@sOrderTableName,
						@sOrderColumnName,
						@iOrderColumnID,
						@iOrderTableID,
						@fOrderAsc;
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @sOrderType = 'F'
						BEGIN
							IF @iLookupColumnID = @iOrderColumnID
							BEGIN
								SET @fLookupColumnDoneF = 1;
								SET @piLookupColumnIndex = @iCount;
							END;
		
							SET @sColumnList = @sColumnList 
								+ CASE
										WHEN LEN(@sColumnList) > 0 THEN ','
										ELSE ''
									END
								+ @sOrderTableName + '.' + @sOrderColumnName;

							SET @iCount = @iCount + 1;
						END
						ELSE
						BEGIN
							SET @sOrderSQL = @sOrderSQL 
								+ CASE
										WHEN LEN(@sOrderSQL) > 0 THEN ','
										ELSE ''
									END
								+ @sOrderTableName + '.' + @sOrderColumnName	
								+CASE
										WHEN @fOrderAsc = 0 THEN ' DESC'
										ELSE ''
									END;
						END;

						IF @iTableID <> @iOrderTableID
						BEGIN
							SET @sTemp = ',' + CONVERT(varchar(max), @iOrderTableID) + ','
							IF CHARINDEX(@sTemp, @sJoinedTables) = 0
							BEGIN
								SET @sJoinedTables = @sJoinedTables + CONVERT(varchar(max), @iOrderTableID) + ',';
								
								SET @sJoinSQL = @sJoinSQL 
									+ ' LEFT OUTER JOIN ' + @sOrderTableName
									+ ' ON ' + @sTableName + '.ID_' + CONVERT(varchar(max), @iOrderTableID)
									+ '=' + @sOrderTableName + '.ID'
							END
						END;

						FETCH NEXT FROM orderCursor INTO 
							@sOrderType, 
							@sOrderTableName,
							@sOrderColumnName,
							@iOrderColumnID,
							@iOrderTableID,
							@fOrderAsc;
					END
					CLOSE orderCursor;
					DEALLOCATE orderCursor;
				
					IF @fLookupColumnDoneF = 0
					BEGIN
						SET @piLookupColumnIndex = @iCount;

						SET @sColumnList = @sColumnList 
							+ CASE
									WHEN LEN(@sColumnList) > 0 THEN ','
									ELSE ''
								END
							+ @sTableName + '.' + @sColumnName;
					END;

					SET @sSelectSQL = 'SELECT ' + @sColumnList + ',';

					IF len(ltrim(rtrim(@sLookupFilterColumnName))) = 0 
					BEGIN
						SET @sSelectSQL = @sSelectSQL
							+ 'null AS [ASRSysLookupFilterValue]';
					END
					ELSE
					BEGIN
						SET @sSelectSQL = @sSelectSQL +
							CASE
								WHEN (@iLookupFilterColumnType = 12) -- Character
									OR (@iLookupFilterColumnType = -1) -- WorkingPattern 
									OR (@iLookupFilterColumnType = -3) THEN -- Photo
									'UPPER(LTRIM(RTRIM(' + @sLookupFilterColumnName + ')))'
								WHEN (@iLookupFilterColumnType = 11) THEN-- Date
									'CASE WHEN ' + @sLookupFilterColumnName + ' IS NULL THEN '''' ELSE CONVERT(varchar(100), ' + @sLookupFilterColumnName + ', 112) END'
								ELSE
									@sLookupFilterColumnName
							END 
							+ ' AS [ASRSysLookupFilterValue]';
					END;

					SET @psDefaultValue = @sDefaultValue;

					SET @sSelectSQL = @sSelectSQL
						+ ' FROM ' + @sTableName 
						+ @sJoinSQL
						+ CASE	
							WHEN len(@sOrderSQL) > 0 THEN ' ORDER BY ' + @sOrderSQL
							ELSE ''
						END;

					EXEC (@sSelectSQL);
				END
			END
