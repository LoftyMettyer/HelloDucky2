CREATE PROCEDURE [dbo].[sp_ASRIntGetAddFromWaitingListRecords] (
	@piTableID 			integer, 
	@piViewID 			integer, 
	@piOrderID 			integer,
	@piCourseRecordID	integer,
	@pfError 			bit 			OUTPUT,
	@piRecordsRequired	integer,
	@pfFirstPage		bit				OUTPUT,
	@pfLastPage			bit				OUTPUT,
	@psLocateValue		varchar(MAX),
	@piColumnType		integer			OUTPUT,
	@psAction			varchar(255),
	@piTotalRecCount	integer			OUTPUT,
	@piFirstRecPos		integer			OUTPUT,
	@piCurrentRecCount	integer,
	@psErrorMessage		varchar(MAX)	OUTPUT,
	@piColumnSize		integer			OUTPUT,
	@piColumnDecimals	integer			OUTPUT
)
AS
BEGIN

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
		@sTempString		varchar(MAX),
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
		@fWhereDone			bit,
		@iWLCourseTitleColumnID	integer,
		@sWLCourseTitleColumnName	sysname,
		@iCourseTitleColumnID	integer,
		@sCourseTitleColumnName	sysname,
		@iWLOverrideColumnID	integer,
		@sWLOverrideColumnName	sysname,
		@iGetCount			integer,
		@sCourseTitle		varchar(MAX),
		@iCourseTableID		integer,
		@iCourseRecordID	integer,
		@iWLTableID			integer,
		@sWLTableName		sysname,
		@sWLRealSource		varchar(255),
		@sCourseSource		sysname,
		@iColSize			integer,
		@iColDecs			integer,
		@sActualUserName	sysname;

	/* Initialise variables. */
	SET @pfError = 0;
	SET @psErrorMessage = '';
	SET @sRealSource = '';
	SET @sSelectSQL = '';
	SET @sOrderSQL = '';
	SET @fSelectDenied = 0;
	SET @sExecString = '';
	SET @sDESCstring = ' DESC';
	SET @fFirstColumnAsc = 1;
	SET @sFirstColCode = '';
	SET @sReverseOrderSQL = '';
	SET @fWhereDone = 0;

	/* Clean the input string parameters. */
	IF len(@psLocateValue) > 0 SET @psLocateValue = replace(@psLocateValue, '''', '''''');

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* Get the current course title. */
	/* Get the Course table id. */
	SELECT @iCourseTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseTable'
	IF @iCourseTableID IS NULL SET @iCourseTableID = 0;

	SELECT @iCourseTitleColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_CourseTitle'
	IF @iCourseTitleColumnID IS NULL SET @iCourseTitleColumnID = 0;
	
	IF @iCourseTitleColumnID > 0 
	BEGIN
		SELECT @sCourseTitleColumnName = columnName
		FROM ASRSysColumns
		WHERE columnID = @iCourseTitleColumnID;
	END
	IF @sCourseTitleColumnName IS NULL SET @sCourseTitleColumnName = '';

	/* Get the @sCourseTitle value for the given course record. */
	DECLARE courseSourceCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT sysobjects.name
	FROM sysprotects
	INNER JOIN sysobjects ON sysprotects.id = sysobjects.id
	INNER JOIN syscolumns ON sysprotects.id = syscolumns.id
	WHERE sysprotects.uid = @iUserGroupID
		AND sysprotects.action = 193 
		AND (sysprotects.protectType = 205 OR sysprotects.protectType = 204)
		AND syscolumns.name = @sCourseTitleColumnName
		AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
			ASRSysTables.tableID = @iCourseTableID 
			UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iCourseTableID)
		AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
		AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
		OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
		AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		
	OPEN courseSourceCursor;
	FETCH NEXT FROM courseSourceCursor INTO @sCourseSource;
	WHILE (@@fetch_status = 0) AND (@sCourseTitle IS null)
	BEGIN
		SET @sTempExecString = 'SELECT @sCourseTitle = ' + @sCourseTitleColumnName + 
			' FROM ' + @sCourseSource +
			' WHERE id = ' + convert(nvarchar(255), @piCourseRecordID);
		SET @sTempParamDefinition = N'@sCourseTitle varchar(MAX) OUTPUT';
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sCourseTitle OUTPUT;

		FETCH NEXT FROM courseSourceCursor INTO @sCourseSource;
	END
	CLOSE courseSourceCursor;
	DEALLOCATE courseSourceCursor;

	IF @sCourseTitle IS null
	BEGIN
		SET @pfError = 1;
		SET @psErrorMessage = 'Unable to read the course title from the current Course record.';
		RETURN		
	END

	/* Get the WL table real source. */
	/* NB. To reach this point we have already checked that the user has 'read' permission on the Training Booking - Status column. */
	SELECT @iWLTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_WaitListTable';
	IF @iWLTableID IS NULL SET @iWLTableID = 0;

	SELECT @sWLTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @iWLTableID;

	SELECT @iWLCourseTitleColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_WaitListCourseTitle';
	IF @iWLCourseTitleColumnID IS NULL SET @iWLCourseTitleColumnID = 0;
	
	IF @iWLCourseTitleColumnID > 0 
	BEGIN
		SELECT @sWLCourseTitleColumnName = columnName
		FROM ASRSysColumns
		WHERE columnID = @iWLCourseTitleColumnID;
	END
	IF @sWLCourseTitleColumnName IS NULL SET @sWLCourseTitleColumnName = '';

	SELECT @iWLOverrideColumnID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_WaitListOverRideColumn';
	IF @iWLOverrideColumnID IS NULL SET @iWLOverrideColumnID = 0;

	IF @iWLOverrideColumnID > 0 
	BEGIN
		SELECT @sWLOverrideColumnName = columnName
		FROM ASRSysColumns
		WHERE columnID = @iWLOverrideColumnID;
	END
	IF @sWLOverrideColumnName IS NULL SET @sWLOverrideColumnName = '';

	SELECT @iChildViewID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @iWLTableID
		AND [role] = @sUserGroupName;
	IF @iChildViewID IS null SET @iChildViewID = 0;
		
	IF @iChildViewID > 0 
	BEGIN
		SET @sWLRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sWLTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_');
		SET @sWLRealSource = left(@sWLRealSource, 255);
	END

	IF @piRecordsRequired <= 0 SET @piRecordsRequired = 1000;
	SET @psAction = UPPER(@psAction);
	IF (@psAction <> 'MOVEPREVIOUS') AND 
		(@psAction <> 'MOVENEXT') AND 
		(@psAction <> 'MOVELAST') AND 
		(@psAction <> 'LOCATE')
	BEGIN
		SET @psAction = 'MOVEFIRST';
	END

	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @piTableID;

	/* Get the real source of the given table/view. */
	IF @iTableType <> 2 /* ie. top-level or lookup */
	BEGIN
		IF @piViewID > 0 
		BEGIN	
			/* RealSource is the view. */	
			SELECT @sRealSource = viewName
			FROM ASRSysViews
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

	IF len(@sRealSource) = 0
	BEGIN
		SET @pfError = 1;
		RETURN;
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
	SELECT DISTINCT c.tableID
	FROM ASRSysOrderItems  oi
		INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
	WHERE oi.orderID = @piOrderID;

	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID;
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
				AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
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
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END

		FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	END
	CLOSE tablesCursor;
	DEALLOCATE tablesCursor;

	/*Get the permissions for the override column if it has been set*/
	IF @iWLOverrideColumnID > 0 
	BEGIN
		INSERT INTO @columnPermissions
		SELECT 
			@iWLTableID,
			@sWLRealSource,
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
			AND sysobjects.name = @sWLRealSource
			AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
			OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0
			AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));

		/* Get the select permission on the column. */
		SELECT @fSelectGranted = selectGranted
		FROM @columnPermissions
		WHERE tableViewName = @sWLRealSource
			AND columnName = @sWLOverrideColumnName;

		IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

		IF @fSelectGranted = 1
		BEGIN
			/* The user DOES have SELECT permission on the column in the current table/view. */
			/* Find column. */
			SET @sTempString = CASE 
					WHEN (len(@sSelectSQL) > 0) THEN ',' 
					ELSE '' 
				END + 
				@sWLRealSource + '.' + @sWLOverrideColumnName;				
			SET @sSelectSQL = @sSelectSQL + @sTempString;

			/* Override Column. */
			IF len(@sOrderSQL) = 0 
			BEGIN
				SET @fAscending = 1;
				SELECT 
					@iDataType = ASRSysColumns.dataType,
					@iColSize = ASRSysColumns.size,
					@iColDecs = ASRSysColumns.decimals
				FROM ASRSysColumns 
				WHERE ASRSysColumns.columnId = @iWLOverrideColumnID;

				SET @piColumnType = @iDataType;
				SET @fFirstColumnAsc = @fAscending;
				SET @sFirstColCode = @sWLRealSource + '.' + @sWLOverrideColumnName;
				SET @piColumnSize = @iColSize;
				SET @piColumnDecimals = @iColDecs;
			END
			SET @sOrderSQL = @sOrderSQL + 
			CASE 
				WHEN len(@sOrderSQL) > 0 THEN ',' 
				ELSE '' 
			END + 
			@sWLRealSource + '.' + @sWLOverrideColumnName +
			CASE 
				WHEN @fAscending = 0 THEN ' DESC' 
				ELSE '' 
			END		
		END			
	END

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT c.tableID,	oi.columnID, c.columnName, t.tableName,	oi.ascending,	oi.type, c.dataType, c.size, c.decimals
	FROM ASRSysOrderItems oi
		INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
		INNER JOIN ASRSysTables t ON t.tableID = c.tableID
	WHERE oi.orderID = @piOrderID
		AND c.dataType <> -4 AND c.datatype <> -3
	ORDER BY oi.sequence;

	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iColSize, @iColDecs;

	/* Check if the order exists. */
	IF  @@fetch_status <> 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END

	WHILE (@@fetch_status = 0)
	BEGIN
		SET @fSelectGranted = 0;

		IF @iColumnTableId = @piTableID
		BEGIN
			/* Base table. */
			/* Get the select permission on the column. */
			SELECT @fSelectGranted = selectGranted
			FROM @columnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName;

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;

			IF @fSelectGranted = 1
			BEGIN
				/* The user DOES have SELECT permission on the column in the current table/view. */
				IF @sType = 'F'
				BEGIN
					/* Find column. */
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName;
					SET @sSelectSQL = @sSelectSQL + @sTempString;

				END
				ELSE
				BEGIN
					/* Order column. */

					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType;
						SET @fFirstColumnAsc = @fAscending;
						SET @sFirstColCode = @sRealSource + '.' + @sColumnName;
						SET @piColumnSize = @iColSize;
						SET @piColumnDecimals = @iColDecs;
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
						END;		
				END
			END
			ELSE
			BEGIN
				/* The user does NOT have SELECT permission on the column in the current table/view. */
				SET @fSelectDenied = 1;
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
				AND columnName = @sColumnName;

			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
	
			IF @fSelectGranted = 1 
			BEGIN
				/* Column COULD be read directly from the parent table. */
				/* The user DOES have SELECT permission on the column in the parent table. */
				IF @sType = 'F'
				BEGIN
					/* Find column. */
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + 
						@sColumnTableName + '.' + @sColumnName;
					SET @sSelectSQL = @sSelectSQL + @sTempString;
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType;
						SET @fFirstColumnAsc = @fAscending;
						SET @sFirstColCode = @sColumnTableName + '.' + @sColumnName;
						SET @piColumnSize = @iColSize;
						SET @piColumnDecimals = @iColDecs;
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
						END;
				END

				/* Add the table to the array of tables/views to join if it has not already been added. */
				SELECT @iTempCount = COUNT(tableViewName)
				FROM @joinParents
				WHERE tableViewName = @sColumnTableName;

				IF @iTempCount = 0
				BEGIN
					INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID);
				END
			END
			ELSE	
			BEGIN
				SET @sSubString = '';

				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT tableViewName
				FROM @columnPermissions
				WHERE tableID = @iColumnTableId
					AND tableViewName <> @sColumnTableName
					AND columnName = @sColumnName
					AND selectGranted = 1;

				OPEN viewCursor;
				FETCH NEXT FROM viewCursor INTO @sViewName;
				WHILE (@@fetch_status = 0)
				BEGIN
					/* Column CAN be read from the view. */
					IF len(@sSubString) = 0 SET @sSubString = 'CASE';

					SET @sSubString = @sSubString +
						' WHEN NOT ' + @sViewName + '.' + @sColumnName + ' IS NULL THEN ' + @sViewName + '.' + @sColumnName;
		
					/* Add the view to the array of tables/views to join if it has not already been added. */
					SELECT @iTempCount = COUNT(tableViewName)
					FROM @joinParents
					WHERE tableViewname = @sViewName;

					IF @iTempCount = 0
					BEGIN
						INSERT INTO @joinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId);
					END

					FETCH NEXT FROM viewCursor INTO @sViewName;
				END
				CLOSE viewCursor;
				DEALLOCATE viewCursor;

				IF len(@sSubString) > 0
				BEGIN
					SET @sSubString = @sSubString +	' ELSE NULL END';

					IF @sType = 'F'
					BEGIN
						/* Find column. */
						SET @sTempString = CASE 
								WHEN (len(@sSelectSQL) > 0) THEN ',' 
								ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END;
						SET @sSelectSQL = @sSelectSQL + @sTempString;
					END
					ELSE
					BEGIN
						/* Order column. */
						IF len(@sOrderSQL) = 0 
						BEGIN
							SET @piColumnType = @iDataType;
							SET @fFirstColumnAsc = @fAscending;
							SET @sFirstColCode = @sSubString;
							SET @piColumnSize = @iColSize;
							SET @piColumnDecimals = @iColDecs;
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
							END;		
					END
				END
				ELSE
				BEGIN
					/* The user does NOT have SELECT permission on the column any of the parent views. */
					SET @fSelectDenied = 1;
				END	
			END
		END

		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iColSize, @iColDecs;
	END
	CLOSE orderCursor;
	DEALLOCATE orderCursor;

	/* Add the ID column to the order string. */
	SET @sOrderSQL = @sOrderSQL + 
		CASE WHEN len(@sOrderSQL) > 0 THEN ',' ELSE '' END + 
		@sRealSource + '.ID';

	/* Create the reverse order string if required. */
	IF (@psAction <> 'MOVEFIRST') 
	BEGIN
		SET @sRemainingSQL = @sOrderSQL;

		SET @iLastCharIndex = 0;
		SET @iCharIndex = CHARINDEX(',', @sOrderSQL);
		WHILE @iCharIndex > 0 
		BEGIN
 			IF UPPER(SUBSTRING(@sOrderSQL, @iCharIndex - LEN(@sDESCstring), LEN(@sDESCstring))) = @sDESCstring
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - LEN(@sDESCstring) - @iLastCharIndex) + ', ';
			END
			ELSE
			BEGIN
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - @iLastCharIndex) + @sDESCstring + ', ';
			END

			SET @iLastCharIndex = @iCharIndex;
			SET @iCharIndex = CHARINDEX(',', @sOrderSQL, @iLastCharIndex + 1);
	
			SET @sRemainingSQL = SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, LEN(@sOrderSQL) - @iLastCharIndex);
		END
		SET @sReverseOrderSQL = @sReverseOrderSQL + @sRemainingSQL + @sDESCstring;
	END

	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource +
		' INNER JOIN ' + @sWLRealSource +
		' ON (' + @sRealSource + '.id = ' + @sWLRealSource + '.id_' + convert(nvarchar(255), @piTableID) +
		' AND ' + @sWLRealSource + '.' + @sWLCourseTitleColumnName + ' = ''' + replace(@sCourseTitle,'''','''''') + ''')';
	SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT;
	SET @piTotalRecCount = @iCount;

	IF (len(@sSelectSQL) > 0)
	BEGIN
		SET @sTempString = ',' + @sRealSource + '.ID';
		SET @sSelectSQL = @sSelectSQL + @sTempString;

		SET @sExecString = 'SELECT ';

		IF @psAction = 'MOVEFIRST' OR @psAction = 'LOCATE'
		BEGIN
			SET @sTempString = 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' ';
			SET @sExecString = @sExecString + @sTempString;
		END
		
		SET @sTempString = @sSelectSQL;
		SET @sExecString = @sExecString + @sTempString;

		SET @sTempString = ' FROM ' + @sWLRealSource;
		SET @sExecString = @sExecString + @sTempString;

		SET @sTempString = ' INNER JOIN ' + @sRealSource +
			' ON (' + @sWLRealSource + '.id_' + convert(nvarchar(255), @piTableID) + '=' + @sRealSource + '.id)' +
			' WHERE ' + @sWLRealSource + '.' + @sWLCourseTitleColumnName + '=''' + replace(@sCourseTitle,'''','''''') + ''''
		SET @sExecString = @sExecString + @sTempString;

		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, tableID
			FROM @joinParents;

		OPEN joinCursor;
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sTempString = ' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';
			SET @sExecString = @sExecString + @sTempString;

			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		END
		CLOSE joinCursor;
		DEALLOCATE joinCursor;

		IF (@psAction = 'MOVELAST')
		BEGIN
			SET @fWhereDone = 1;
			
			SET @sTempString = ' AND ' + @sWLRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sWLRealSource + '.ID' +
				' FROM ' + @sWLRealSource;
			SET @sExecString = @sExecString + @sTempString;

			SET @sTempString = ' INNER JOIN ' + @sRealSource +
				' ON (' + @sWLRealSource + '.id_' + convert(nvarchar(255), @piTableID) + '=' + @sRealSource + '.id)' +
				' WHERE ' + @sWLRealSource + '.' + @sWLCourseTitleColumnName + '=''' + replace(@sCourseTitle,'''','''''') + '''';
			SET @sExecString = @sExecString + @sTempString;

		END

		IF @psAction = 'MOVENEXT' 
		BEGIN
			SET @fWhereDone = 1
			IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
			BEGIN
				SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1);
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired;
			END

			SET @sTempString = ' AND ' + @sWLRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sWLRealSource + '.ID' +
				' FROM ' + @sWLRealSource;
			SET @sExecString = @sExecString + @sTempString;
				
			SET @sTempString = ' INNER JOIN ' + @sRealSource +
				' ON (' + @sWLRealSource + '.id_' + convert(nvarchar(255), @piTableID) + '=' + @sRealSource + '.id)' +
				' WHERE ' + @sWLRealSource + '.' + @sWLCourseTitleColumnName + '=''' + replace(@sCourseTitle,'''','''''') + '''';
			SET @sExecString = @sExecString + @sTempString;

			SET @sTempString = ' AND ' + @sWLRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired - 1) + ' ' + @sWLRealSource + '.ID' +
				' FROM ' + @sWLRealSource;
			SET @sExecString = @sExecString + @sTempString;

			SET @sTempString = ' INNER JOIN ' + @sRealSource +
				' ON (' + @sWLRealSource + '.id_' + convert(nvarchar(255), @piTableID) + '=' + @sRealSource + '.id)' +
				' WHERE ' + @sWLRealSource + '.' + @sWLCourseTitleColumnName + '=''' + replace(@sCourseTitle,'''','''''') + '''';
			SET @sExecString = @sExecString + @sTempString;

		END

		IF @psAction = 'MOVEPREVIOUS'
		BEGIN
			SET @fWhereDone = 1;
			IF @piFirstRecPos <= @piRecordsRequired
			BEGIN
				SET @iGetCount = @piFirstRecPos - 1;
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired;		
			END

			SET @sTempString = ' AND ' + @sWLRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sWLRealSource + '.ID' +
				' FROM ' + @sWLRealSource;
			SET @sExecString = @sExecString + @sTempString;
				
			SET @sTempString = ' INNER JOIN ' + @sRealSource +
				' ON (' + @sWLRealSource + '.id_' + convert(nvarchar(255), @piTableID) + '=' + @sRealSource + '.id)' +
				' WHERE ' + @sWLRealSource + '.' + @sWLCourseTitleColumnName + '=''' + replace(@sCourseTitle,'''','''''') + '''';
			SET @sExecString = @sExecString + @sTempString;

			SET @sTempString = ' AND ' + @sWLRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sWLRealSource + '.ID' +
				' FROM ' + @sWLRealSource;
			SET @sExecString = @sExecString + @sTempString;

			SET @sTempString = ' INNER JOIN ' + @sRealSource +
				' ON (' + @sWLRealSource + '.id_' + convert(nvarchar(255), @piTableID) + '=' + @sRealSource + '.id)' +
				' WHERE ' + @sWLRealSource + '.' + @sWLCourseTitleColumnName + '=''' + replace(@sCourseTitle,'''','''''') + '''';
			SET @sExecString = @sExecString + @sTempString;
		END

		/* Add the filter code. */

		IF @psAction = 'MOVENEXT' OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sTempString = ' ORDER BY ' + @sOrderSQL + ')';
			SET @sExecString = @sExecString + @sTempString;
		END

		IF (@psAction = 'MOVELAST') OR (@psAction = 'MOVENEXT') OR (@psAction = 'MOVEPREVIOUS')
		BEGIN
			SET @sTempString = ' ORDER BY ' + @sReverseOrderSQL + ')';
			SET @sExecString = @sExecString + @sTempString;
		END

		IF (@psAction = 'LOCATE')
		BEGIN
			SET @fWhereDone = 1;
			SET @sLocateCode = ' AND (' + @sFirstColCode;

			IF (@piColumnType = 12) OR (@piColumnType = -1) /* Character or Working Pattern column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + '''';

					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL';
					END
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + 
						@sFirstColCode + ' LIKE ''' + @psLocateValue + '%'' OR ' + @sFirstColCode + ' IS NULL';
				END

			END

			IF @piColumnType = 11 /* Date column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NOT NULL  OR ' + @sFirstColCode + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + '''';
					END
				END
				ELSE
				BEGIN
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + @sFirstColCode + ' IS NULL';
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
						END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END;
				END
			END

			IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
			BEGIN
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + @psLocateValue;

					IF convert(float, @psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL';
					END
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + @psLocateValue + ' OR ' + @sFirstColCode + ' IS NULL';
				END
			END

			SET @sLocateCode = @sLocateCode + ')';
			SET @sTempString = @sLocateCode;
			SET @sExecString = @sExecString + @sTempString;

		END

		/* Add the ORDER BY code to the find record selection string if required. */
		SET @sTempString = ' ORDER BY ' + @sOrderSQL;
		SET @sExecString = @sExecString + @sTempString;

	END

	/* Set the IsFirstPage, IsLastPage flags, and the page number. */
	IF @psAction = 'MOVEFIRST'
	BEGIN
		SET @piFirstRecPos = 1;
		SET @pfFirstPage = 1;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount <= @piRecordsRequired THEN 1
				ELSE 0
			END;
	END
	IF @psAction = 'MOVENEXT'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos + @piCurrentRecCount;
		SET @pfFirstPage = 0;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END;
	END
	IF @psAction = 'MOVEPREVIOUS'
	BEGIN
		SET @piFirstRecPos = @piFirstRecPos - @iGetCount;
		IF @piFirstRecPos <= 0 SET @piFirstRecPos = 1;
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount < @piFirstRecPos + @iGetCount THEN 1
				ELSE 0
			END;
	END
	IF @psAction = 'MOVELAST'
	BEGIN
		SET @piFirstRecPos = @piTotalRecCount - @piRecordsRequired + 1;
		IF @piFirstRecPos < 1 SET @piFirstRecPos = 1;
		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END;
		SET @pfLastPage = 1;
	END

	IF @psAction = 'LOCATE'
	BEGIN
		SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource + 
			' INNER JOIN ' + @sWLRealSource +
			' ON (' + @sRealSource + '.id = ' + @sWLRealSource + '.id_' + convert(nvarchar(255), @piTableID) +
			' AND ' + @sWLRealSource + '.' + @sWLCourseTitleColumnName + ' = ''' + replace(@sCourseTitle,'''','''''') + ''')';
			
		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, tableID
		FROM @joinParents;

		OPEN joinCursor;
		FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		WHILE (@@fetch_status = 0)
		BEGIN
			SET @sTempExecString = @sTempExecString + 
				' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';

			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
		END
		CLOSE joinCursor;
		DEALLOCATE joinCursor;

		SET @sTempExecString = @sTempExecString + @sLocateCode;

		SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
		EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iTemp OUTPUT;

		IF @iTemp <=0 

		BEGIN
			SET @piFirstRecPos = @piTotalRecCount + 1;
		END
		ELSE
		BEGIN
			SET @piFirstRecPos = @piTotalRecCount - @iTemp + 1;
		END

		SET @pfFirstPage = 
			CASE 
				WHEN @piFirstRecPos = 1 THEN 1
				ELSE 0
			END;
		SET @pfLastPage = 
			CASE 
				WHEN @piTotalRecCount <= @piFirstRecPos + @piRecordsRequired THEN 1
				ELSE 0
			END;
	END

	/* Return a recordset of the required columns in the required order from the given table/view. */
	IF (len(@sExecString) > 0)
	BEGIN
		EXECUTE sp_executeSQL @sExecString;
	END
END