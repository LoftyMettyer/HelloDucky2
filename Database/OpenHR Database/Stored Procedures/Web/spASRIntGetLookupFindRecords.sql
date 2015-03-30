CREATE PROCEDURE [dbo].[spASRIntGetLookupFindRecords] (
	@piLookupColumnID 	integer,
	@piRecordsRequired	integer,
	@pfFirstPage		bit			OUTPUT,
	@pfLastPage			bit			OUTPUT,
	@psLocateValue		varchar(MAX),
	@piColumnType		integer		OUTPUT,
	@piColumnSize		integer		OUTPUT,
	@piColumnDecimals	integer		OUTPUT,
	@psAction			varchar(255),
	@piTotalRecCount	integer		OUTPUT,
	@piFirstRecPos		integer		OUTPUT,
	@piCurrentRecCount	integer,
	@psFilterValue		varchar(MAX),
	@piCallingColumnID	integer,
	@pfOverrideFilter	bit
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the lookup find records, given the table and column IDs.
		@piTableID = the ID of the table on which the find is based.
		@piLookupColumnID = the ID of the column on which the find is based.
	NB. No permissions need to be read, as all users have read permission on lookup tables.
	*/
	DECLARE @sTableName		sysname,
		@iTableID 			integer, 
		@sColumnName 		sysname,
		@sColumnName2 		sysname,
		@iOrderID			integer,
		@sSelectSQL			varchar(MAX),
		@sOrderSQL 			varchar(MAX),
		@sFilterValuesSQL	varchar(MAX),
		@sExecString		nvarchar(MAX),
		@iTemp				integer,
		@sRemainingSQL		varchar(MAX),
		@iLastCharIndex		integer,
		@iCharIndex 		integer,
		@sTempExecString	nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@sLocateCode		varchar(MAX),
		@sReverseOrderSQL 	varchar(MAX),
		@iCount				integer,
		@iGetCount			integer,
		@sColumnTemp		sysname,
		@iLookupFilterColumnID	integer,
		@iLookupFilterOperator	integer,
		@iLookupFilterColumnDataType	integer;

	/* Initialise variables. */
	SET @sSelectSQL = ''
	SET @sOrderSQL = ''
	SET @sFilterValuesSQL = ''
	SET @sExecString = ''
	SET @sReverseOrderSQL = ''

	/* Clean the input string parameters. */
	IF len(@psFilterValue) > 0 SET @psFilterValue = replace(@psFilterValue, '''', '''''')
	IF len(@psLocateValue) > 0 SET @psLocateValue = replace(@psLocateValue, '''', '''''')
	
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
	SELECT @sColumnName = ASRSysColumns.columnName,
		@iTableID = ASRSysColumns.tableID,
		@piColumnType = ASRSysColumns.dataType,
		@piColumnSize = ASRSysColumns.size,
		@piColumnDecimals = ASRSysColumns.decimals
	FROM ASRSysColumns
	WHERE ASRSysColumns.columnId = @piLookupColumnID

	/* Get the table name and default order. */
	SELECT @sTableName = ASRSysTables.tableName,
		@iOrderID = ASRSysTables.defaultOrderID
	FROM ASRSysTables
	WHERE ASRSysTables.tableID = @iTableID

	SET @sSelectSQL = @sTableName + '.' + @sColumnName
	SET @sOrderSQL = @sTableName + '.' + @sColumnName

	/* Filter the values if required */
	SELECT @iLookupFilterColumnID  = ASRSysColumns.LookupFilterColumnID,
		@iLookupFilterOperator = ASRSysColumns.LookupFilterOperator
	FROM ASRSysColumns
	WHERE ASRSysColumns.columnId = @piCallingColumnID

	IF (@iLookupFilterColumnID > 0) and (@pfOverrideFilter = 0)
	BEGIN
		SELECT @sColumnTemp = ASRSysColumns.columnName,
			@iLookupFilterColumnDataType = ASRSysColumns.dataType
		FROM ASRSysColumns
		WHERE ASRSysColumns.columnId = @iLookupFilterColumnID

		IF @iLookupFilterColumnDataType = -7 /* Boolean */
		BEGIN
			SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' = ' 
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
					SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' = ' + @psFilterValue + ') '
					IF convert(float, @psFilterValue) = 0
					BEGIN
						SET @sFilterValuesSQL = @sFilterValuesSQL +
							' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
					END
				END

				IF @iLookupFilterOperator = 2 /* NOT Equal To */
				BEGIN
					SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' <> ' + @psFilterValue + ') '
					IF convert(float, @psFilterValue) = 0
					BEGIN
						SET @sFilterValuesSQL = @sFilterValuesSQL +
							' AND (' + @sTableName + '.' + @sColumnTemp  + ' IS NOT null) '
					END
				END

				IF @iLookupFilterOperator = 3 /* Is At Most */
				BEGIN
					SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' <= ' + @psFilterValue + ') '
					IF convert(float, @psFilterValue) >= 0
					BEGIN
						SET @sFilterValuesSQL = @sFilterValuesSQL +
							' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
					END
				END

				IF @iLookupFilterOperator = 4 /* Is At Least */
				BEGIN
					SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' >= ' + @psFilterValue + ') '
					IF convert(float, @psFilterValue) <= 0
					BEGIN
						SET @sFilterValuesSQL = @sFilterValuesSQL +
							' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
					END
				END

				IF @iLookupFilterOperator = 5 /* Is More Than */
				BEGIN
					SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' > ' + @psFilterValue + ') '
					IF convert(float, @psFilterValue) < 0
					BEGIN
						SET @sFilterValuesSQL = @sFilterValuesSQL +
							' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
					END
				END

				IF @iLookupFilterOperator = 6 /* Is Less Than */
				BEGIN
					SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' < ' + @psFilterValue + ') '
					IF convert(float, @psFilterValue) > 0
					BEGIN
						SET @sFilterValuesSQL = @sFilterValuesSQL +
							' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
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
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' = ''' + @psFilterValue + ''') '
						END
						ELSE
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 8 /* NOT On */
					BEGIN
						IF len(@psFilterValue) = 10
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' <> ''' + @psFilterValue + ''') ' +
								' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
						END
						ELSE
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' IS NOT null) '
						END
					END

					IF @iLookupFilterOperator = 12 /* On OR Before*/
					BEGIN
						IF len(@psFilterValue) = 10
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' <= ''' + @psFilterValue + ''') ' +
								' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
						END
						ELSE
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
						END
					END

					IF @iLookupFilterOperator = 11 /* On OR After*/
					BEGIN
						IF len(@psFilterValue) = 10
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' >= ''' + @psFilterValue + ''') ' 
						END
						ELSE
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' IS null)' +
								' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS NOT null)'
						END
					END

					IF @iLookupFilterOperator = 9 /* After*/
					BEGIN
						IF len(@psFilterValue) = 10
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' > ''' + @psFilterValue + ''') ' 
						END
						ELSE
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' IS NOT null)'
						END
					END

					IF @iLookupFilterOperator = 10 /* Before*/
					BEGIN
						IF len(@psFilterValue) = 10
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' < ''' + @psFilterValue + ''') ' +
								' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
						END
						ELSE
						BEGIN
							SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' IS null)' +
								' AND (' + @sTableName + '.' + @sColumnTemp  + ' IS NOT null)'
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
								SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' = '''') ' +
									' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS null) '
							END
							ELSE
							BEGIN
								SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' = ''' + @psFilterValue + ''') '
							END
						END

						IF @iLookupFilterOperator = 16 /* Is NOT*/
						BEGIN
							IF len(@psFilterValue) = 0
							BEGIN
								SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' <> '''') ' +
									' AND (' + @sTableName + '.' + @sColumnTemp  + ' IS NOT null) '
							END
							ELSE
							BEGIN
								SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' <> ''' + @psFilterValue + ''') '
							END
						END

						IF @iLookupFilterOperator = 13 /* Contains*/
						BEGIN
							IF len(@psFilterValue) = 0
							BEGIN
								SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' IS null) ' +
									' OR (' + @sTableName + '.' + @sColumnTemp  + ' IS NOT null) '
							END
							ELSE
							BEGIN
								SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' LIKE ''%' + @psFilterValue + '%'') '
							END
						END

						IF @iLookupFilterOperator = 15 /* Does NOT Contain*/
						BEGIN
							IF len(@psFilterValue) = 0
							BEGIN
								SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' IS null) ' +
									' AND (' + @sTableName + '.' + @sColumnTemp  + ' IS NOT null) '
							END
							ELSE
							BEGIN
								SET @sFilterValuesSQL = '(' + @sTableName + '.' + @sColumnTemp  + ' NOT LIKE ''%' + @psFilterValue + '%'') '
							END
						END
					END
				END
			END
		END
	END

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT c.columnName
	FROM ASRSysOrderItems oi
	INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
	INNER JOIN ASRSysTables t ON t.tableID = c.tableID
	WHERE oi.orderID = @iOrderID AND oi.type = 'F'
		AND oi.columnID <> @piLookupColumnID
		AND c.dataType <> -4 AND c.datatype <> -3
	ORDER BY oi.sequence;

	OPEN orderCursor
	FETCH NEXT FROM orderCursor INTO @sColumnName2
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @sSelectSQL = @sSelectSQL +  ','  + @sTableName + '.' + @sColumnName2

		FETCH NEXT FROM orderCursor INTO @sColumnName2
	END
	CLOSE orderCursor
	DEALLOCATE orderCursor

	/* Create the reverse order string if required. */
	IF (@psAction <> 'MOVEFIRST') 
	BEGIN
		SET @sReverseOrderSQL = @sTableName + '.' + @sColumnName + ' DESC, ' + @sTableName + '.ID DESC'
	END

	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(*) ' +
												 'FROM (SELECT DISTINCT ' + @sSelectSQL +
															' FROM ' + @sTableName 
	IF len(@sFilterValuesSQL) > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterValuesSQL
	SET @sTempExecString = @sTempExecString	+ ') ' + 'distinctTable'
	SET @sTempParamDefinition = N'@recordCount integer OUTPUT'
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT
	SET @piTotalRecCount = @iCount

	SET @sExecString = 'SELECT DISTINCT ' 

	IF @psAction = 'MOVEFIRST' OR @psAction = 'LOCATE' 
	BEGIN
		SET @sExecString = @sExecString + 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' '
	END
	SET @sExecString = @sExecString + @sSelectSQL +
		' FROM ' + @sTableName

	IF (@psAction = 'MOVEFIRST') AND LEN(@sFilterValuesSQL) > 0
	BEGIN
		SET @sExecString = @sExecString + ' WHERE ' + @sFilterValuesSQL
	END

	IF (@psAction = 'MOVELAST') 
	BEGIN
		IF LEN(@sFilterValuesSQL) > 0 SET @sFilterValuesSQL = @sFilterValuesSQL + ' AND '

		SET @sExecString = @sExecString + 
			' WHERE '  + @sFilterValuesSQL + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName
	END
	IF @psAction = 'MOVENEXT' 
	BEGIN
		IF LEN(@sFilterValuesSQL) > 0 SET @sFilterValuesSQL = @sFilterValuesSQL + ' AND '

		IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
		BEGIN
			SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1)
		END
		ELSE
		BEGIN
			SET @iGetCount = @piRecordsRequired
		END
		SET @sExecString = @sExecString + 
			' WHERE ' + @sFilterValuesSQL + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName

		SET @sExecString = @sExecString + 
			' WHERE ' + @sFilterValuesSQL + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired  - 1) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName
	END
	IF @psAction = 'MOVEPREVIOUS'
	BEGIN
		IF LEN(@sFilterValuesSQL) > 0 SET @sFilterValuesSQL = @sFilterValuesSQL + ' AND '

		IF @piFirstRecPos <= @piRecordsRequired
		BEGIN
			SET @iGetCount = @piFirstRecPos - 1
		END
		ELSE
		BEGIN
			SET @iGetCount = @piRecordsRequired
		END
		SET @sExecString = @sExecString + 
			' WHERE ' + @sFilterValuesSQL + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName

		SET @sExecString = @sExecString + 
			' WHERE ' + @sFilterValuesSQL + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName
	END

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
		IF LEN(@sFilterValuesSQL) > 0
			SET @sLocateCode = ' WHERE ((' +@sFilterValuesSQL + ') AND ' +@sTableName + '.' + @sColumnName 
		ELSE 
			SET @sLocateCode = ' WHERE (' + @sTableName + '.' + @sColumnName

		IF (@piColumnType = 12) OR (@piColumnType = -1) /* Character or Working Pattern column */
		BEGIN
			SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + ''''

			IF len(@psLocateValue) = 0
			BEGIN
				SET @sLocateCode = @sLocateCode + ' OR ' + @sTableName + '.' + @sColumnName + ' IS NULL'
			END
		END

		IF @piColumnType = 11 /* Date column */
		BEGIN
			IF len(@psLocateValue) = 0
			BEGIN
				SET @sLocateCode = @sLocateCode + ' IS NOT NULL  OR ' + @sTableName + '.' + @sColumnName + ' IS NULL'
			END
			ELSE
			BEGIN
				SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + ''''
			END
		END

		IF @piColumnType = -7 /* Logic column */
		BEGIN
			SET @sLocateCode = @sLocateCode + ' >= ' + 
				CASE
					WHEN @psLocateValue = 'True' THEN '1'
					ELSE '0'
				END
		END

		IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
		BEGIN
			SET @sLocateCode = @sLocateCode + ' >= ' + @psLocateValue

			IF convert(float, @psLocateValue) = 0
			BEGIN
				SET @sLocateCode = @sLocateCode + ' OR ' + @sTableName + '.' + @sColumnName + ' IS NULL'
			END
		END

		SET @sLocateCode = @sLocateCode + ')'
		SET @sExecString = @sExecString + @sLocateCode
	END

	/* Add the ORDER BY code to the find record selection string if required. */
	SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL

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
		SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sTableName + '.id) FROM ' + @sTableName + @sLocateCode
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
	EXECUTE sp_executeSQL @sExecString;
END