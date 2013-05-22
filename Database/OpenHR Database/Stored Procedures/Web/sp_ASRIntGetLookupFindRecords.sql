CREATE PROCEDURE [dbo].[sp_ASRIntGetLookupFindRecords] (
	@piColumnID 		integer,
	@piRecordsRequired	integer,
	@pfFirstPage		bit			OUTPUT,
	@pfLastPage			bit			OUTPUT,
	@psLocateValue		varchar(MAX),
	@piColumnType		integer		OUTPUT,
	@piColumnSize		integer		OUTPUT,
	@piColumnDecimals	integer		OUTPUT,
	@psAction			varchar(100),
	@piTotalRecCount	integer		OUTPUT,
	@piFirstRecPos		integer		OUTPUT,
	@piCurrentRecCount	integer
)
AS
BEGIN
	/* Return a recordset of the lookup find records, given the table and column IDs.
		@piTableID = the ID of the table on which the find is based.
		@piColumnID = the ID of the column on which the find is based.
	NB. No permissions need to be read, as all users have read permission on lookup tables.
	*/
	DECLARE @sTableName		sysname,
		@iTableID 			integer, 
		@sColumnName 		sysname,
		@sColumnName2 		sysname,
		@iOrderID			integer,
		@sSelectSQL			varchar(MAX),
		@sOrderSQL 			varchar(MAX),
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
		@iGetCount			integer;

	/* Initialise variables. */
	SET @sSelectSQL = '';
	SET @sOrderSQL = '';
	SET @sExecString = '';
	SET @sReverseOrderSQL = '';

	IF @piRecordsRequired <= 0 SET @piRecordsRequired = 1000;
	SET @psAction = UPPER(@psAction);
	IF (@psAction <> 'MOVEPREVIOUS') AND 
		(@psAction <> 'MOVENEXT') AND 
		(@psAction <> 'MOVELAST') AND 
		(@psAction <> 'LOCATE')
	BEGIN
		SET @psAction = 'MOVEFIRST';
	END

	/* Get the column name. */
	SELECT @sColumnName = ASRSysColumns.columnName,
		@iTableID = ASRSysColumns.tableID,
		@piColumnType = ASRSysColumns.dataType,
		@piColumnSize = ASRSysColumns.size,
		@piColumnDecimals = ASRSysColumns.decimals
	FROM [dbo].[ASRSysColumns]
	WHERE ASRSysColumns.columnID = @piColumnID;

	/* Get the table name and default order. */
	SELECT @sTableName = ASRSysTables.tableName,
		@iOrderID = ASRSysTables.defaultOrderID
	FROM [dbo].[ASRSysTables]
	WHERE ASRSysTables.tableID = @iTableID;

	SET @sSelectSQL = @sTableName + '.' + @sColumnName;
	SET @sOrderSQL = @sTableName + '.' + @sColumnName + ', ' + @sTableName + '.ID';

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.columnName
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @iOrderID
		AND ASRSysOrderItems.type = 'F'
		AND ASRSysOrderItems.columnID <> @piColumnID
	ORDER BY ASRSysOrderItems.sequence;

	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @sColumnName2;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @sSelectSQL = @sSelectSQL +  ','  + @sTableName + '.' + @sColumnName2;

		FETCH NEXT FROM orderCursor INTO @sColumnName2;
	END
	CLOSE orderCursor;
	DEALLOCATE orderCursor;

	/* Create the reverse order string if required. */
	IF (@psAction <> 'MOVEFIRST') 
	BEGIN
		SET @sReverseOrderSQL = @sTableName + '.' + @sColumnName + ' DESC, ' + @sTableName + '.ID DESC';
	END

	/* Get the total number of records. */
	SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sTableName + '.id) FROM ' + @sTableName;
	SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT;
	SET @piTotalRecCount = @iCount;

	SET @sExecString = 'SELECT ';

	IF @psAction = 'MOVEFIRST' OR @psAction = 'LOCATE' 
	BEGIN
		SET @sExecString = @sExecString + 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' ';
	END
	SET @sExecString = @sExecString + @sSelectSQL + ' FROM ' + @sTableName;

	IF (@psAction = 'MOVELAST') 
	BEGIN
		SET @sExecString = @sExecString + 
			' WHERE ' + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName;
	END
	
	IF @psAction = 'MOVENEXT' 
	BEGIN
		IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
		BEGIN
			SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1);
		END
		ELSE
		BEGIN
			SET @iGetCount = @piRecordsRequired;
		END
		SET @sExecString = @sExecString + 
			' WHERE ' + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName;

		SET @sExecString = @sExecString + 
			' WHERE ' + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(8000), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired  - 1) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName;
			
	END
	IF @psAction = 'MOVEPREVIOUS'
	BEGIN
		IF @piFirstRecPos <= @piRecordsRequired
		BEGIN
			SET @iGetCount = @piFirstRecPos - 1;
		END
		ELSE
		BEGIN
			SET @iGetCount = @piRecordsRequired;
		END
		SET @sExecString = @sExecString + 
			' WHERE ' + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName;

		SET @sExecString = @sExecString + 
			' WHERE ' + @sTableName + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sTableName + '.ID' +
			' FROM ' + @sTableName;
	END

	IF @psAction = 'MOVENEXT' OR (@psAction = 'MOVEPREVIOUS')
	BEGIN
		SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL + ')';
	END
	IF (@psAction = 'MOVELAST') OR (@psAction = 'MOVENEXT') OR (@psAction = 'MOVEPREVIOUS')
	BEGIN
		SET @sExecString = @sExecString + ' ORDER BY ' + @sReverseOrderSQL + ')';
	END

	IF (@psAction = 'LOCATE')
	BEGIN
		SET @sLocateCode = ' WHERE (' + @sTableName + '.' + @sColumnName;

		IF (@piColumnType = 12) OR (@piColumnType = -1) /* Character or Working Pattern column */
		BEGIN
			SET @sLocateCode = @sLocateCode + ' >= ''' + replace(@psLocateValue, '''', '''''') + '''';

			IF len(@psLocateValue) = 0
			BEGIN
				SET @sLocateCode = @sLocateCode + ' OR ' + @sTableName + '.' + @sColumnName + ' IS NULL';
			END
		END

		IF @piColumnType = 11 /* Date column */
		BEGIN
			IF len(@psLocateValue) = 0
			BEGIN
				SET @sLocateCode = @sLocateCode + ' IS NOT NULL  OR ' + @sTableName + '.' + @sColumnName + ' IS NULL';
			END
			ELSE
			BEGIN
				SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + '''';
			END
		END

		IF @piColumnType = -7 /* Logic column */
		BEGIN
			SET @sLocateCode = @sLocateCode + ' >= ' + 
				CASE
					WHEN @psLocateValue = 'True' THEN '1'
					ELSE '0'
				END;
		END

		IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
		BEGIN
			SET @sLocateCode = @sLocateCode + ' >= ' + @psLocateValue;

			IF convert(float, @psLocateValue) = 0
			BEGIN
				SET @sLocateCode = @sLocateCode + ' OR ' + @sTableName + '.' + @sColumnName + ' IS NULL';
			END
		END

		SET @sLocateCode = @sLocateCode + ')';
		SET @sExecString = @sExecString + @sLocateCode;
	END

	/* Add the ORDER BY code to the find record selection string if required. */
	SET @sExecString = @sExecString + ' ORDER BY ' + @sOrderSQL;

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
			END
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
		SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sTableName + '.id) FROM ' + @sTableName + @sLocateCode;
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
				WHEN @piTotalRecCount < @piFirstRecPos + @piRecordsRequired THEN 1
				ELSE 0
			END;
	END

	-- Return a recordset of the required columns in the required order from the given table/view.
	EXECUTE sp_executeSQL @sExecString;
	
END