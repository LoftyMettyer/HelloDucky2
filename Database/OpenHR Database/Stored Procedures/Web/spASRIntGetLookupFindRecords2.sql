CREATE PROCEDURE [dbo].[spASRIntGetLookupFindRecords2] (
	@piTableID 			integer, 
	@piViewID 			integer, 
	@piOrderID 			integer,
	@piLookupColumnID 	integer,
	@piRecordsRequired	integer,
	@pfFirstPage		bit		OUTPUT,
	@pfLastPage			bit		OUTPUT,
	@psLocateValue		varchar(MAX),
	@piColumnType		integer		OUTPUT,
	@piColumnSize		integer		OUTPUT,
	@piColumnDecimals	integer		OUTPUT,
	@psAction			varchar(MAX),
	@piTotalRecCount	integer		OUTPUT,
	@piFirstRecPos		integer		OUTPUT,
	@piCurrentRecCount	integer,
	@psFilterValue		varchar(MAX),
	@piCallingColumnID	integer,
	@piLookupColumnGridNumber	integer		OUTPUT,
	@pfOverrideFilter	bit
)
AS
BEGIN
	/* Return a recordset of the link find records for the current user, given the table/view and order IDs.
		@piTableID = the ID of the table on which the find is based.
		@piViewID = the ID of the view on which the find is based.
		@piOrderID = the ID of the order we are using.
		@pfError = 1 if errors occured in getting the find records. Else 0.
	*/
	
	SET NOCOUNT ON;
	
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iTableType			integer,
		@sTableName			sysname,
		@sRealSource 		sysname,
		@iChildViewID 		integer,
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@iColumnID 			integer,
		@sColumnName 		sysname,
		@sColumnTableName 	sysname,
		@fAscending 		bit,
		@sType	 			varchar(10),
		@iDataType 			integer,
		@fSelectGranted 	bit,
		@sSelectSQL			varchar(MAX),
		@sOrderSQL 			varchar(MAX),
		@sExecString		nvarchar(MAX),
		@fSelectDenied		bit,
		@iTempCount 		integer,
		@sSubString			varchar(MAX),
		@sViewName 			varchar(255),
		@sTableViewName 	sysname,
		@iJoinTableID 		integer,
		@iTemp				integer,
		@sRemainingSQL		varchar(MAX),
		@iLastCharIndex		integer,
		@iCharIndex 		integer,
		@sDESCstring		varchar(5),
		@sTempExecString	nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@fFirstColumnAsc	bit,
		@sFirstColCode		varchar(MAX),
		@sLocateCode		varchar(MAX),
		@sReverseOrderSQL 	varchar(MAX),
		@iCount				integer,
		@iGetCount			integer,
		@iColSize			integer,
		@iColDecs			integer,
		@fLookupColumnDone	bit,
		@sLookupColumnName	sysname,
		@iLookupTableID		integer,
		@iLookupColumnType	integer,
		@iLookupColumnSize	integer,
		@iLookupColumnDecimals integer,
		@iCount2			integer,
		@sFilterSQL			nvarchar(MAX),
		@sColumnTemp		sysname,
		@iLookupFilterColumnID	integer,
		@iLookupFilterOperator	integer,
		@iLookupFilterColumnDataType	integer,
		@sActualUserName	sysname;

	/* Initialise variables. */
	SET @sRealSource = ''
	SET @sSelectSQL = ''
	SET @sOrderSQL = ''
	SET @fSelectDenied = 0
	SET @sExecString = ''
	SET @sDESCstring = ' DESC'
	SET @fFirstColumnAsc = 1
	SET @sFirstColCode = ''
	SET @sReverseOrderSQL = ''
	SET @fLookupColumnDone = 0
	SET @piLookupColumnGridNumber = 0
	SET @sFilterSQL = ''

	/* Clean the input string parameters. */
	IF len(@psFilterValue) > 0 SET @psFilterValue = replace(@psFilterValue, '''', '''''')
	IF len(@psLocateValue) > 0 SET @psLocateValue = replace(@psLocateValue, '''', '''''')
	
	/* Get the current user's group ID. */
	EXEC spASRIntGetActualUserDetails
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT

	IF @piRecordsRequired <= 0 SET @piRecordsRequired = 1000
	SET @psAction = UPPER(@psAction)
	IF (@psAction <> 'MOVEPREVIOUS') AND 
		(@psAction <> 'MOVENEXT') AND 
		(@psAction <> 'MOVELAST') AND 
		(@psAction <> 'LOCATE')
	BEGIN
		SET @psAction = 'MOVEFIRST'
	END

	/* Get the column name. */
	SELECT @sLookupColumnName = ASRSysColumns.columnName,
		@iLookupTableID = ASRSysColumns.tableID,
		@iLookupColumnType = ASRSysColumns.dataType,
		@iLookupColumnSize = ASRSysColumns.size,
		@iLookupColumnDecimals = ASRSysColumns.decimals
	FROM ASRSysColumns
	WHERE ASRSysColumns.columnId = @piLookupColumnID

	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @piTableID

	/* Get the real source of the given table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		IF @piViewID > 0 
		BEGIN	
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM ASRSysViews
			WHERE viewID = @piViewID
		END
		ELSE
		BEGIN
			SET @sRealSource = @sTableName
		END 
	END
	ELSE
	BEGIN
		SELECT @iChildViewID = childViewID
		FROM ASRSysChildViews2
		WHERE tableID = @piTableID
			AND role = @sUserGroupName
			
		IF @iChildViewID IS null SET @iChildViewID = 0
			
		IF @iChildViewID > 0 
		BEGIN
			SET @sRealSource = 'ASRSysCV' + 
				convert(varchar(1000), @iChildViewID) +
				'#' + replace(@sTableName, ' ', '_') +
				'#' + replace(@sUserGroupName, ' ', '_')
			SET @sRealSource = left(@sRealSource, 255)
		END
	END

	IF len(@sRealSource) = 0
	BEGIN
		RETURN
	END

	/* Create a temporary table to hold the tables/views that need to be joined. */
	DECLARE @joinParents TABLE(
		tableViewName	sysname,
		tableID			integer);

	/* Create a temporary table of the 'select' column permissions for all tables/views used in the order. */
	DECLARE @columnPermissions TABLE(
		tableID			integer,
		tableViewName	sysname,
		columnName		sysname,
		selectGranted	bit);

	/* Loop through the tables used in the order, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT DISTINCT ASRSysColumns.tableID
	FROM ASRSysOrderItems 
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	WHERE ASRSysOrderItems.orderID = @piOrderID

	OPEN tablesCursor
	FETCH NEXT FROM tablesCursor INTO @iTempTableID
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @piTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @columnPermissions
			SELECT 
				@iTempTableID,
				@sRealSource,
				syscolumns.name,
				CASE protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM sysprotects
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.action = 193 
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			INSERT INTO @columnPermissions
			SELECT 
				@iTempTableID,
				sysobjects.name,
				syscolumns.name,
				CASE protectType
				        	WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM sysprotects
			INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
			INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
			WHERE sysprotects.uid = @iUserGroupID
				AND sysprotects.action = 193 
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
					ASRSysTables.tableID = @iTempTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID
	END
	CLOSE tablesCursor
	DEALLOCATE tablesCursor

	/* Create the lookup filter string. NB. We already know that the user has SELECT permission on this from the spASRIntGetLookupViews stored procedure.*/
	SELECT @iLookupFilterColumnID = ASRSysColumns.LookupFilterColumnID,
		@iLookupFilterOperator = ASRSysColumns.LookupFilterOperator
	FROM ASRSysColumns
	WHERE ASRSysColumns.columnId = @piCallingColumnID

	IF (@iLookupFilterColumnID > 0) and (@pfOverrideFilter = 0)
	BEGIN
		SELECT @sColumnTemp = ASRSysColumns.columnName,
			@iLookupFilterColumnDataType = ASRSysColumns.dataType
		FROM ASRSysColumns
		WHERE ASRSysColumns.columnId = @iLookupFilterColumnID

		SELECT @iCount = COUNT(*)
		FROM @columnPermissions
		WHERE columnName = @sColumnTemp
			AND selectGranted = 1

		IF @iCount > 0 AND @psFilterValue <> ''
		BEGIN
			IF @iLookupFilterColumnDataType = -7 /* Boolean */
			BEGIN
				SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = '
					+ CASE
						WHEN UPPER(@psFilterValue) = 'TRUE' THEN '1'
						WHEN UPPER(@psFilterValue) = 'FALSE' THEN '0'
						ELSE @psFilterValue
					END
					+ ') '
			END
			ELSE
			BEGIN
				IF (@iLookupFilterColumnDataType = 2) OR (@iLookupFilterColumnDataType = 4) /* Numeric, Integer */
				BEGIN
					IF @iLookupFilterOperator = 1 /* Equals */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) = 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 2 /* NOT Equal To */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <> ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) = 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' AND (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
						END
					END

					IF @iLookupFilterOperator = 3 /* Is At Most */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <= ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) >= 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 4 /* Is At Least */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' >= ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) <= 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 5 /* Is More Than */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' > ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) < 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 6 /* Is Less Than */
					BEGIN
						SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' < ' + @psFilterValue + ') '
						IF convert(float, @psFilterValue) > 0
						BEGIN
							SET @sFilterSQL = @sFilterSQL +
								' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
						END
					END
				END
				ELSE
				BEGIN
					IF (@iLookupFilterColumnDataType = 11) /* Date */
					BEGIN
						IF @iLookupFilterOperator = 7 /* On */
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = ''' + @psFilterValue + ''') '
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
						END

						IF @iLookupFilterOperator = 8 /* NOT On */
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <> ''' + @psFilterValue + ''') ' +
									' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
							END
						END

						IF @iLookupFilterOperator = 12 /* On OR Before*/
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <= ''' + @psFilterValue + ''') ' +
									' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
						END

						IF @iLookupFilterOperator = 11 /* On OR After*/
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' >= ''' + @psFilterValue + ''') ' 
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null)' +
									' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null)'
							END
						END

						IF @iLookupFilterOperator = 9 /* After*/
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' > ''' + @psFilterValue + ''') ' 
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null)'
							END
						END

						IF @iLookupFilterOperator = 10 /* Before*/
						BEGIN
							IF len(@psFilterValue) = 10
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' < ''' + @psFilterValue + ''') ' +
									' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
							END
							ELSE
							BEGIN
								SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null)' +
									' AND (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null)'
							END
						END
					END
					ELSE
					BEGIN
						IF (@iLookupFilterColumnDataType = 12) OR (@iLookupFilterColumnDataType = -3) OR (@iLookupFilterColumnDataType = -1) /* varchar, working patter, photo*/
						BEGIN
							IF @iLookupFilterOperator = 14 /* Is */
							BEGIN
								IF len(@psFilterValue) = 0
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = '''') ' +
										' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS null) '
								END
								ELSE
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' = ''' + @psFilterValue + ''') '
								END
							END

							IF @iLookupFilterOperator = 16 /* Is NOT*/
							BEGIN
								IF len(@psFilterValue) = 0
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <> '''') ' +
										' AND (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
								END
								ELSE
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' <> ''' + @psFilterValue + ''') '
								END
							END

							IF @iLookupFilterOperator = 13 /* Contains*/
							BEGIN
								IF len(@psFilterValue) = 0
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null) ' +
										' OR (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
								END
								ELSE
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' LIKE ''%' + @psFilterValue + '%'') '
								END
							END

							IF @iLookupFilterOperator = 15 /* Does NOT Contain*/
							BEGIN
								IF len(@psFilterValue) = 0
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' IS null) ' +
										' AND (' + @sRealSource + '.' + @sColumnTemp  + ' IS NOT null) '
								END
								ELSE
								BEGIN
									SET @sFilterSQL = '(' + @sRealSource + '.' + @sColumnTemp  + ' NOT LIKE ''%' + @psFilterValue + '%'') '
								END
							END
						END
					END
				END
			END
		END
	END

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 

	SELECT ASRSysColumns.tableID,
		ASRSysOrderItems.columnID, 
		ASRSysColumns.columnName,
	    	ASRSysTables.tableName,
		ASRSysOrderItems.ascending,
		ASRSysOrderItems.type,
		ASRSysColumns.dataType,
		ASRSysColumns.size,
		ASRSysColumns.decimals
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnId
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @piOrderID
	ORDER BY ASRSysOrderItems.sequence

	OPEN orderCursor
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iColSize, @iColDecs

	/* Check if the order exists. */
	IF  @@fetch_status <> 0
	BEGIN
		RETURN
	END
	SET @iCount2 = 0

	WHILE (@@fetch_status = 0) OR (@fLookupColumnDone = 0)
	BEGIN
		SET @fSelectGranted = 0

		IF (@@fetch_status <> 0)
		BEGIN
			SET @iColumnTableId = @iLookupTableID
			SET @iColumnId = @piLookupColumnID
			SET @sColumnName = @sLookupColumnName
			SET @sColumnTableName = @sTableName
			SET @fAscending = 1
			SET @sType = 'F'
			SET @iDataType = @iLookupColumnType
			SET @iColSize = @iLookupColumnSize
			SET @iColDecs = @iLookupColumnDecimals
		END

		IF (@iColumnId  = @piLookupColumnID ) AND (@sType = 'F')
		BEGIN
			SET @fLookupColumnDone = 1
			SET @piLookupColumnGridNumber = @iCount2
		END

		IF @iColumnTableId = @piTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = selectGranted
			FROM @columnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0

			IF @fSelectGranted = 1
			BEGIN
				/* The user DOES have SELECT permission on the column in the current table/view. */
				IF @sType = 'F'
				BEGIN

					/* Find column. */
					SET @sSelectSQL = @sSelectSQL + 
						CASE 
							WHEN len(@sSelectSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName
					SET @iCount2 = @iCount2 + 1
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType
						SET @piColumnSize = @iColSize
						SET @piColumnDecimals = @iColDecs
						SET @fFirstColumnAsc = @fAscending
						SET @sFirstColCode = @sRealSource + '.' + @sColumnName
					END

					SET @sOrderSQL = @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName +
						CASE 
							WHEN @fAscending = 0 THEN ' DESC' 
							ELSE '' 
						END				
				END
			END
			ELSE
			BEGIN
				/* The user does NOT have SELECT permission on the column in the current table/view. */
				SET @fSelectDenied = 1
			END	
		END
		ELSE
		BEGIN
			/* Parent of the base table. */
			/* Get the select permission on the column. */

			/* Check if the column is selectable directly from the table. */
			SELECT @fSelectGranted = selectGranted
			FROM @columnPermissions
			WHERE tableID = @iColumnTableId
				AND tableViewName = @sColumnTableName
				AND columnName = @sColumnName

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				/* The user DOES have SELECT permission on the column in the parent table. */
				IF @sType = 'F'
				BEGIN
					/* Find column. */
					SET @sSelectSQL = @sSelectSQL + 
						CASE 
							WHEN len(@sSelectSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sColumnTableName + '.' + @sColumnName
					SET @iCount2 = @iCount2 + 1
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType
						SET @piColumnSize = @iColSize
						SET @piColumnDecimals = @iColDecs
						SET @fFirstColumnAsc = @fAscending
						SET @sFirstColCode = @sColumnTableName + '.' + @sColumnName
					END

					SET @sOrderSQL = @sOrderSQL + 
						CASE 
							WHEN len(@sOrderSQL) > 0 THEN ',' 
							ELSE '' 
						END + 
						@sColumnTableName + '.' + @sColumnName + 
						CASE 
							WHEN @fAscending = 0 THEN ' DESC' 
							ELSE '' 
						END				
				END

				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)

				FROM @joinParents
				WHERE tableViewName = @sColumnTableName

				IF @iTempCount = 0
				BEGIN
					INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID)
				END
			END
			ELSE	
			BEGIN
				SET @sSubString = ''

				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @columnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND selectGranted = 1

				OPEN viewCursor
				FETCH NEXT FROM viewCursor INTO @sViewName
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					IF len(@sSubString) = 0 SET @sSubString = 'CASE'

					SET @sSubString = @sSubString +
						' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName 
		
					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @joinParents
					WHERE tableViewname = @sViewName

					IF @iTempCount = 0
					BEGIN
						INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId)
					END

					FETCH NEXT FROM viewCursor INTO @sViewName
				END
				CLOSE viewCursor
				DEALLOCATE viewCursor

				IF len(@sSubString) > 0
				BEGIN
					SET @sSubString = @sSubString +
						' ELSE NULL END'

					IF @sType = 'F'
					BEGIN
						/* Find column. */
						SET @sSubString = @sSubString +
							' AS [' + @sColumnName + ']'

						SET @sSelectSQL = @sSelectSQL + 
							CASE 
								WHEN len(@sSelectSQL) > 0 THEN ',' 
								ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END
						SET @iCount2 = @iCount2 + 1
					END
					ELSE
					BEGIN
						/* Order column. */
						IF len(@sOrderSQL) = 0 
						BEGIN
							SET @piColumnType = @iDataType
							SET @piColumnSize = @iColSize
							SET @piColumnDecimals = @iColDecs
							SET @fFirstColumnAsc = @fAscending
							SET @sFirstColCode = @sSubString
						END

						SET @sOrderSQL = @sOrderSQL + 
							CASE 
								WHEN len(@sOrderSQL) > 0 THEN ',' 
								ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END + 
							CASE 
								WHEN @fAscending = 0 THEN ' DESC' 
								ELSE '' 
							END				
					END
				END
				ELSE
				BEGIN
					/* The user does NOT have SELECT permission on the column any of the parent views. */
					SET @fSelectDenied = 1
				END	
			END
		END

		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iColSize, @iColDecs
	END
	CLOSE orderCursor
	DEALLOCATE orderCursor

	/* Add the ID column to the order string. */
	SET @sOrderSQL = @sOrderSQL + 
		CASE WHEN len(@sOrderSQL) > 0 THEN ',' ELSE '' END + 
		@sRealSource + '.ID'

	/* Create the reverse order string if required. */
	IF (@psAction <> 'MOVEFIRST') 
	BEGIN
		SET @sRemainingSQL = @sOrderSQL

		SET @iLastCharIndex = 0
		SET @iCharIndex = CHARINDEX(',', @sOrderSQL)
		WHILE @iCharIndex > 0 
		BEGIN
 			IF UPPER(SUBSTRING(@sOrderSQL, @iCharIndex - LEN(@sDESCstring), LEN(@sDESCstring))) = @sDESCstring
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - LEN(@sDESCstring) - @iLastCharIndex) + ', '
			END
			ELSE
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - @iLastCharIndex) + @sDESCstring + ', '
			END

			SET @iLastCharIndex = @iCharIndex
			SET @iCharIndex = CHARINDEX(',', @sOrderSQL, @iLastCharIndex + 1)
	
			SET @sRemainingSQL = SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, LEN(@sOrderSQL) - @iLastCharIndex)
		END
		SET @sReverseOrderSQL = @sReverseOrderSQL + @sRemainingSQL + @sDESCstring

	END

	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource

	IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL

	SET @sTempParamDefinition = N'@recordCount integer OUTPUT'
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT
	SET @piTotalRecCount = @iCount

	IF len(@sSelectSQL) > 0 
	BEGIN
		SET @sSelectSQL = @sSelectSQL + ',' + @sRealSource + '.ID'
		SET @sExecString = 'SELECT ' 

		IF @psAction = 'MOVEFIRST' OR @psAction = 'LOCATE' 
		BEGIN
			SET @sExecString = @sExecString + 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' '
		END
		SET @sExecString = @sExecString + @sSelectSQL + 
			' FROM ' + @sRealSource

		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @joinParents

		OPEN joinCursor
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sExecString = @sExecString + 
				' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID'

			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		END
		CLOSE joinCursor
		DEALLOCATE joinCursor

		IF (@psAction = 'MOVELAST')
		BEGIN
			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
		END

		IF @psAction = 'MOVENEXT' 
		BEGIN
			IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
			BEGIN
				SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1)
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired
			END
			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource

			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired  - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
		END
		IF @psAction = 'MOVEPREVIOUS'
		BEGIN
			IF @piFirstRecPos <= @piRecordsRequired
			BEGIN
				SET @iGetCount = @piFirstRecPos - 1
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired
			END
			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource

			SET @sExecString = @sExecString + 
				' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
		END

		IF len(@sFilterSQL) > 0 SET @sExecString = @sExecString + ' WHERE ' + @sFilterSQL

		IF @psAction = 'MOVENEXT' OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL + ')'
		END

		IF (@psAction = 'MOVELAST') OR (@psAction = 'MOVENEXT') OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sExecString = @sExecString + ' ORDER BY ' + @sReverseOrderSQL + ')'
		END

		IF (@psAction = 'LOCATE')
		BEGIN
			IF len(@sFilterSQL) > 0 
			BEGIN
				SET @sLocateCode = ' AND (' + @sFirstColCode 
			END
			ELSE
			BEGIN
				SET @sLocateCode = ' WHERE (' + @sFirstColCode 
			END

			IF (@piColumnType = 12) OR (@piColumnType = -1) /* Character or Working Pattern column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + ''''

					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL'
					END
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + 
						@sFirstColCode + ' LIKE ''' + @psLocateValue + '%'' OR ' + @sFirstColCode + ' IS NULL'
				END

			END

			IF @piColumnType = 11 /* Date column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NOT NULL  OR ' + @sFirstColCode + ' IS NULL'
					END
					ELSE
					BEGIN

						SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + ''''
					END
				END
				ELSE
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NULL'
					END
					ELSE
					BEGIN
						SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + @sFirstColCode + ' IS NULL'
					END
				END
			END

			IF @piColumnType = -7 /* Logic column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END
				END
			END

			IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + @psLocateValue

					IF convert(float, @psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL'
					END
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + @psLocateValue + ' OR ' + @sFirstColCode + ' IS NULL'
				END

			END

			SET @sLocateCode = @sLocateCode + ')'
			SET @sExecString = @sExecString + @sLocateCode
		END

		/* Add the ORDER BY code to the find record selection string if required. */
		SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL
	END

	/* Set the IsFirstPage, IsLastPage flags, and the page number. */
	IF @psAction = 'MOVEFIRST'
	BEGIN
		SET @piFirstRecPos = 1
		SET @pfFirstPage = 1
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount <= @piRecordsRequired THEN 1
				ELSE 0
			END
	END
	IF @psAction = 'MOVENEXT'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos + @piCurrentRecCount
		SET @pfFirstPage = 0
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END
	END
	IF @psAction = 'MOVEPREVIOUS'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos - @iGetCount
		IF @piFirstRecPos <= 0 SET @piFirstRecPos = 1
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END
	END
	IF @psAction = 'MOVELAST'
	BEGIN
		SET @piFirstRecPos = @piTotalRecCount - @piRecordsRequired + 1
		IF @piFirstRecPos < 1 SET @piFirstRecPos = 1
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END
		SET @pfLastPage = 1
	END

	IF @psAction = 'LOCATE'
	BEGIN
		SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource

		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @joinParents

		OPEN joinCursor
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sTempExecString = @sTempExecString + 
				' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID'

			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID
		END
		CLOSE joinCursor
		DEALLOCATE joinCursor

		IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL

		SET @sTempExecString = @sTempExecString + @sLocateCode

		SET @sTempParamDefinition = N'@recordCount integer OUTPUT'
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iTemp OUTPUT

		IF @iTemp <=0 
		BEGIN
			SET @piFirstRecPos = @piTotalRecCount + 1
		END
		ELSE
		BEGIN
			SET @piFirstRecPos = @piTotalRecCount - @iTemp + 1
		END

		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @piRecordsRequired THEN 1
				ELSE 0
			END
	END

	/* Return a recordset of the required columns in the required order from the given table/view. */
	IF (len(@sExecString) > 0)
	BEGIN
		EXEC sp_executeSQL @sExecString;
	END
END
GO

