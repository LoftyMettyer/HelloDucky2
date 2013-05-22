CREATE PROCEDURE [dbo].[spASRIntGetLookupFilterValue] (
	@piScreenID			integer,
	@piColumnID			integer,
	@piTableID			integer,
	@piViewID			integer,
	@piRecordID			integer,
	@psFilterValue		varchar(MAX)	OUTPUT,
	@piParentTableID	integer,
	@piParentRecordID	integer,
	@pfError 			bit	OUTPUT
)
AS
BEGIN
	DECLARE
		@iLookupFilterValueID	integer,
		@sRealSource			sysname,
		@sLookupFilterValueName	sysname,
		@iCount					integer,
		@iTableType				integer,
		@sTableName				sysname,
		@iChildViewID			integer,
		@sUserGroupName			sysname,
		@iUserGroupID			integer,
		@sSQL					nvarchar(MAX),
		@sSQLParam				nvarchar(500),
		@dblResult				float,
		@fResult				bit,
		@dtResult				datetime,
		@sResult				varchar(MAX),
		@iDataType				integer,
		@sActualUserName		sysname,
		@iDfltValueExprID		integer,
		@sDefaultValue			varchar(MAX),
		@lngParentTableID		integer;

	SET @psFilterValue = '';
	SET @pfError = 0;

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
					
	SET @psFilterValue = '';

	SELECT @iLookupFilterValueID = lookupFilterValueID
	FROM [dbo].[ASRSysColumns]
	WHERE columnID = @piColumnID;

	IF @iLookupFilterValueID IS null SET @iLookupFilterValueID = 0;

	IF @iLookupFilterValueID > 0 
	BEGIN
		/* Check if the looup filter value column is in the screen. If so, we don't need to find it again. */
		SELECT @iCount = COUNT(*)
		FROM [dbo].[ASRSysControls]
		WHERE screenID = @piScreenID
			AND columnId = @iLookupFilterValueID;

		IF @iCount = 0 
		BEGIN
			IF @piRecordID = 0
			BEGIN
				/* New record. */
				SELECT @iDfltValueExprID = dfltValueExprID, 
					@sDefaultValue = defaultValue, 
					@iDataType = dataType
				FROM [dbo].[ASRSysColumns]
				WHERE columnID = @iLookupFilterValueID;

				IF @iDfltValueExprID IS null SET @iDfltValueExprID = 0;
				IF @sDefaultValue IS null SET @sDefaultValue = '';
				IF @iDataType IS null SET @iDataType = 0;

				IF @iDfltValueExprID > 0
				BEGIN
					/* Calculated value as the default. */
					SELECT @iCount = COUNT(*)
					FROM sysobjects
					WHERE id = object_id(N'sp_ASRDfltExpr_' + convert(sysname, @iDfltValueExprID))
					AND OBJECTPROPERTY(id, N'IsProcedure') = 1;

					IF @iCount > 0
					BEGIN
						SET @sSQL = 'EXEC sp_ASRDfltExpr_' + convert(nvarchar(100), @iDfltValueExprID) + ' @result OUTPUT';

						DECLARE parents_cursor CURSOR LOCAL FAST_FORWARD FOR 
							SELECT parentID
							FROM ASRSysRelations
							WHERE childID = @piTableID
							ORDER BY parentID;

						OPEN parents_cursor;
						FETCH NEXT FROM parents_cursor INTO @lngParentTableID;
						WHILE (@@fetch_status = 0)
						BEGIN
						    IF @lngParentTableID = @piParentTableID
							BEGIN
								SET @sSQL = @sSQL + ',' + convert(nvarchar(100), @piParentRecordID);
							END
							ELSE
							BEGIN
								SET @sSQL = @sSQL + ',0' ;
							END

							FETCH NEXT FROM parents_cursor INTO @lngParentTableID;
						END
						CLOSE parents_cursor;
						DEALLOCATE parents_cursor;

						IF @iDataType = -7 /* Boolean */
						BEGIN
							SET @sSQLParam = N'@result integer OUTPUT';
							EXEC sp_executesql @sSQL, @sSQLParam, @fResult OUTPUT;

							SET @psFilterValue = 
								CASE 
									WHEN @fResult = 0 THEN 'False'
									ELSE 'True'
								END;
						END

						IF (@iDataType = 2) OR (@iDataType = 4) /* Numeric, Integer */
						BEGIN
							SET @sSQLParam = N'@result float OUTPUT';
							EXEC sp_executesql @sSQL, @sSQLParam, @dblResult OUTPUT;

							SET @psFilterValue = convert(varchar(MAX), @dblResult);
						END

						IF (@iDataType = 11) /* Date */
						BEGIN
							SET @sSQLParam = N'@result datetime OUTPUT';
							EXEC sp_executesql @sSQL, @sSQLParam, @dtResult OUTPUT

							SET @psFilterValue = convert(varchar(MAX), @dtResult, 101);
						END

						IF (@iDataType = 12) OR (@iDataType = -3) OR (@iDataType = -1) /* varchar, working pattern, photo*/
						BEGIN
							SET @sSQLParam = N'@result varchar(MAX) OUTPUT';
							EXEC sp_executesql @sSQL, @sSQLParam, @sResult OUTPUT;

							SET @psFilterValue = @sResult;
						END
					END
				END
				ELSE
				BEGIN
					SET @psFilterValue = @sDefaultValue;
				END
			END
			ELSE
			BEGIN
				SELECT @sTableName = tableName,
					@iTableType = tableType
				FROM [dbo].[ASRSysTables]
				WHERE tableID = @piTableID;

				IF @iTableType <> 2 /* ie. top-level or lookup */
				BEGIN
					IF @piViewID > 0 
					BEGIN	
						/* RealSource is the view. */	
						SELECT @sRealSource = viewName
						FROM [dbo].[ASRSysViews]
						WHERE viewID = @piViewID;
					END
					ELSE
					BEGIN
						SET @sRealSource = @sTableName;
					END 
				END
				ELSE
				BEGIN
					SELECT @iChildViewID = childViewID
					FROM ASRSysChildViews2
					WHERE tableID = @piTableID
						AND [role] = @sUserGroupName;
						
					IF @iChildViewID IS null SET @iChildViewID = 0;
						
					IF @iChildViewID > 0 
					BEGIN
						SET @sRealSource = 'ASRSysCV' + 
							convert(varchar(1000), @iChildViewID) +
							'#' + replace(@sTableName, ' ', '_') +
							'#' + replace(@sUserGroupName, ' ', '_');
						SET @sRealSource = left(@sRealSource, 255);
					END
				END
	
				SELECT @sLookupFilterValueName = columnName,
					@iDataType = dataType
				FROM [dbo].[ASRSysColumns]
				WHERE columnID = @iLookupFilterValueID;
	
				/* Check the filter column can be read from the given table/view. */
				SELECT @iCount = COUNT(*)
				FROM sysprotects
				INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
				INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
				WHERE sysprotects.uid = @iUserGroupID
					AND sysprotects.action = 193 
					AND ((sysprotects.protectType = 204) OR (sysprotects.protectType = 205))
					AND syscolumns.name = @sLookupFilterValueName
					AND sysobjects.name = @sRealSource
					AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
					OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
					AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));

				IF @iCount > 0 
				BEGIN
					/* Can read the column from the realSource */
					SET @sSQL = 'SELECT @result = ' + @sLookupFilterValueName 
						+ ' FROM ' + @sRealSource
						+ ' WHERE id = ' + convert(nvarchar(100), @piRecordID);

					IF @iDataType = -7 /* Boolean */
					BEGIN
						SET @sSQLParam = N'@result integer OUTPUT';
						EXEC sp_executesql @sSQL, @sSQLParam, @fResult OUTPUT;

						SET @psFilterValue = 
							CASE 
								WHEN @fResult = 0 or @fResult is null THEN 'False'
								ELSE 'True'
							END;
					END

					IF (@iDataType = 2) OR (@iDataType = 4) /* Numeric, Integer */
					BEGIN
						SET @sSQLParam = N'@result float OUTPUT';
						EXEC sp_executesql @sSQL, @sSQLParam, @dblResult OUTPUT;

						SET @psFilterValue = convert(varchar(8000), @dblResult);
					END

					IF (@iDataType = 11) /* Date */
					BEGIN
						SET @sSQLParam = N'@result datetime OUTPUT';
						EXEC sp_executesql @sSQL, @sSQLParam, @dtResult OUTPUT;

						SET @psFilterValue = convert(varchar(8000), @dtResult, 101);
					END

					IF (@iDataType = 12) OR (@iDataType = -3) OR (@iDataType = -1) /* varchar, working patter, photo*/
					BEGIN
						SET @sSQLParam = N'@result varchar(MAX) OUTPUT';
						EXEC sp_executesql @sSQL, @sSQLParam, @sResult OUTPUT;

						SET @psFilterValue = @sResult;
					END
				END
				ELSE
				/* Column cannot be read from view */
				SET @pfError = 1;
			END
		END
	END
END