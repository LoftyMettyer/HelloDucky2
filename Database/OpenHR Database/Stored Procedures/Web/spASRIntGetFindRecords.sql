CREATE PROCEDURE [dbo].[spASRIntGetFindRecords] (
	@pfError 				bit 			OUTPUT, 
	@pfSomeSelectable 		bit 			OUTPUT, 
	@pfSomeNotSelectable 	bit 			OUTPUT, 
	@psRealSource			varchar(255)	OUTPUT,
	@pfInsertGranted		bit				OUTPUT,
	@pfDeleteGranted		bit				OUTPUT,
	@piTableID 				integer, 
	@piViewID 				integer, 
	@piOrderID 				integer, 
	@piParentTableID		integer,
	@piParentRecordID		integer,
	@psFilterDef			varchar(MAX),
	@piRecordsRequired		integer,
	@pfFirstPage			bit				OUTPUT,
	@pfLastPage				bit				OUTPUT,
	@psLocateValue			varchar(MAX),
	@piColumnType			integer			OUTPUT,
	@piColumnSize			integer			OUTPUT,
	@piColumnDecimals		integer			OUTPUT,
	@psAction				varchar(255),
	@piTotalRecCount		integer			OUTPUT,
	@piFirstRecPos			integer			OUTPUT,
	@piCurrentRecCount		integer,
	@psDecimalSeparator		varchar(255),
	@psLocaleDateFormat		varchar(255),
	@RecordID				integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the find records for the current user, given the table/view and order IDs.
		@pfError = 1 if errors occured in getting the find records. Else 0.
		@pfSomeSelectable = 1 if some find columns were selectable. Else 0.
		@pfSomeNotSelectable = 1 if some find columns were NOT selectable. Else 0.
		@piTableID = the ID of the table on which the find is based.
		@piViewID = the ID of the view on which the find is based.
		@piOrderID = the ID of the order we are using.
		@piParentTableID = the ID of the parent table.
		@piParentRecordID = the ID of the associated record in the parent table.
	*/
	DECLARE @iUserGroupID	integer,
		@sUserGroupName		sysname,
		@iTableType			integer,
		@sTableName			sysname,
		@sRealSource 		sysname,
		@iChildViewID 		integer,
		@iTempTableID 		integer,
		@iColumnTableID 	integer,
		@sColumnName 		sysname,
		@sColumnTableName 	sysname,
		@fAscending 		bit,
		@sType	 			varchar(10),
		@fSelectGranted 	bit,
		@sSelectSQL			varchar(MAX),
		@sOrderSQL 			varchar(MAX),
		@sReverseOrderSQL 	varchar(MAX),
		@fSelectDenied		bit,
		@iTempCount 		integer,
		@sSubString			varchar(MAX),
		@sViewName 			varchar(255),
		@sExecString		nvarchar(MAX),
		@sTempString		varchar(MAX),
		@sTableViewName 	sysname,
		@iJoinTableID 		integer,
		@iDataType 			integer,
		@iTempAction		integer,
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
		@iCount				integer,
		@iGetCount			integer,
		@iSize				integer,
		@iDecimals			integer,
		@bUse1000Separator	bit,
		@sActualLoginName	varchar(250),
		@iIndex1			integer,
		@iIndex2			integer,
		@iIndex3			integer,
		@iColumnID			integer,
		@iOperatorID		integer,
		@sValue				varchar(MAX),
		@sFilterSQL			nvarchar(MAX),
		@sTempLocateValue	varchar(MAX),
		@sSubFilterSQL		nvarchar(MAX),
		@bBlankIfZero		bit,
		@sTempFilterString	varchar(MAX),
		@sJoinSQL			nvarchar(max),
		@psOriginalAction		varchar(255),
		@sThousandColumns		varchar(255),
		@sBlankIfZeroColumns	varchar(255),
		@sTableOrViewName varchar(255);

	DECLARE @FindDefinition TABLE(tableID integer, columnID integer, columnName nvarchar(255), tableName nvarchar(255)
									, ascending bit, type varchar(1), datatype integer, controltype integer, size integer, decimals integer, Use1000Separator bit, BlankIfZero bit, Editable bit
									, LookupTableID integer, LookupColumnID integer, LookupFilterColumnID integer, LookupFilterValueID integer, SpinnerMinimum smallint, SpinnerMaximum smallint, SpinnerIncrement smallint
									, DefaultValue varchar(max), Mask varchar(max), DefaultValueExprID int
									)

	DECLARE @OriginalColumns TABLE(columnID integer, columnName nvarchar(255))

	/* Clean the input string parameters. */
	IF len(@psLocateValue) > 0 SET @psLocateValue = replace(@psLocateValue, '''', '''''');
	/* Initialise variables. */
	SET @pfError = 0;
	SET @pfSomeSelectable = 0;
	SET @pfSomeNotSelectable = 0;
	SET @sDESCstring = ' DESC';
	SET @fFirstColumnAsc = 1;
	SET @sFirstColCode = '';
	SET @piColumnSize = 0;
	SET @piColumnDecimals = 0;
	SET @bUse1000Separator = 0;
	SET @bBlankIfZero = 0;
	SET @sThousandColumns = '';
	SET @sTableOrViewName = '';
	SET @sBlankIfZeroColumns = '';

	SET @sRealSource = '';
	SET @sSelectSQL = '';
	SET @sOrderSQL = '';
	SET @sReverseOrderSQL = '';
	SET @fSelectDenied = 0;
	SET @sExecString = '';
	SET @sFilterSQL = '';
	SET @sTempLocateValue = '';
	
	SET @psOriginalAction = @psAction;
	
	IF @piRecordsRequired <= 0 SET @piRecordsRequired = 1000;
	SET @psAction = UPPER(@psAction);
	IF (@psAction <> 'MOVEPREVIOUS') AND 
		(@psAction <> 'MOVENEXT') AND 
		(@psAction <> 'MOVELAST') AND 
		(@psAction <> 'LOCATE') AND 
		(@psAction <> 'LOCATEID')
	BEGIN
		SET @psAction = 'MOVEFIRST';
	END
	
	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualLoginName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;
		
	-- Create a temporary view of the sysprotects table
	DECLARE @SysProtects TABLE([ID] int, Columns varbinary(8000)
								, [Action] tinyint
								, ProtectType tinyint);
	INSERT INTO @SysProtects
	SELECT p.[ID], p.[Columns], p.[Action], p.ProtectType FROM ASRSysProtectsCache p WHERE p.UID = @iUserGroupID;
	
	-- Create a temporary table to hold the tables/views that need to be joined.
	DECLARE @JoinParents TABLE(tableViewName sysname, tableID integer);

	-- Create a temporary table of the 'select' column permissions for all tables/views used in the order.
	DECLARE @ColumnPermissions TABLE(tableID integer,
								tableViewName	sysname,
								columnName	sysname,
								selectGranted	bit,
								updateGranted	bit);
								
	/* Get the table type and name. */
	SELECT @iTableType = ASRSysTables.tableType,
		@sTableName = ASRSysTables.tableName
	FROM [dbo].[ASRSysTables]
		WHERE ASRSysTables.tableID = @piTableID;
	
	IF (@sTableName IS NULL) 
	BEGIN 
		SET @pfError = 1;
		RETURN;
	END
	/* Get the real source of the given table/view. */
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
		FROM [dbo].[ASRSysChildViews2]
		WHERE tableID = @piTableID
			AND role = @sUserGroupName;
			
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
	SET @psRealSource = @sRealSource;
	
	IF len(@sRealSource) = 0
	BEGIN
		SET @pfError = 1;
		RETURN;
	END

	IF len(@psFilterDef)> 0 
	BEGIN
		WHILE charindex('	', @psFilterDef) > 0
		BEGIN
			SET @sSubFilterSQL = '';
			SET @iIndex1 = charindex('	', @psFilterDef);
			SET @iIndex2 = charindex('	', @psFilterDef, @iIndex1+1);
			SET @iIndex3 = charindex('	', @psFilterDef, @iIndex2+1);
				
			SET @iColumnID = convert(integer, LEFT(@psFilterDef, @iIndex1-1));
			SET @iOperatorID = convert(integer, SUBSTRING(@psFilterDef, @iIndex1+1, @iIndex2-@iIndex1-1));
			SET @sValue = SUBSTRING(@psFilterDef, @iIndex2+1, @iIndex3-@iIndex2-1);
			
			SET @psFilterDef = SUBSTRING(@psFilterDef, @iIndex3+1, LEN(@psFilterDef) - @iIndex3);
			SELECT @iDataType = dataType,
				@sColumnName = columnName
			FROM [dbo].[ASRSysColumns]
			WHERE columnID = @iColumnID;
							
			SET @sColumnName = @sRealSource + '.' + @sColumnName;
			IF (@iDataType = -7) 
			BEGIN
				/* Logic column (must be the equals operator).	*/
				SET @sSubFilterSQL = @sColumnName + ' = ';
			
				IF UPPER(@sValue) = 'TRUE'
				BEGIN
					SET @sSubFilterSQL = @sSubFilterSQL + '1';
				END
				ELSE
				BEGIN
					SET @sSubFilterSQL = @sSubFilterSQL + '0';
				END
			END
			
			IF ((@iDataType = 2) OR (@iDataType = 4)) 
			BEGIN
				/* Numeric/Integer column. */
				/* Replace the locale decimal separator with '.' for SQL's benefit. */
				SET @sValue = REPLACE(@sValue, @psDecimalSeparator, '.');
				IF (@iOperatorID = 1) 
				BEGIN
					/* Equals. */
					SET @sSubFilterSQL = @sColumnName + ' = ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END

				IF (@iOperatorID = 2)
				BEGIN
					/* Not Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' <> ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' AND ' + @sColumnName + ' IS NOT NULL';
					END
				END
				
				IF (@iOperatorID = 3) 
				BEGIN
					/* Less than or Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' <= ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
        
				IF (@iOperatorID = 4) 
				BEGIN
					/* Greater than or Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' >= ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
				IF (@iOperatorID = 5) 
				BEGIN
					/* Greater than. */
					SET @sSubFilterSQL = @sColumnName + ' > ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
				IF (@iOperatorID = 6) 
				BEGIN
					/* Less than.*/
					SET @sSubFilterSQL = @sColumnName + ' < ' + @sValue;
					IF convert(float, @sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
					END
				END
			END
			IF (@iDataType = 11) 
			BEGIN
				/* Date column. */
				IF LEN(@sValue) > 0
				BEGIN
					/* Convert the locale date into the SQL format. */
					/* Note that the locale date has already been validated and formatted to match the locale format. */
					SET @iIndex1 = CHARINDEX('mm', @psLocaleDateFormat);
					SET @iIndex2 = CHARINDEX('dd', @psLocaleDateFormat);
					SET @iIndex3 = CHARINDEX('yyyy', @psLocaleDateFormat);
						
					SET @sValue = SUBSTRING(@sValue, @iIndex1, 2) + '/' 
						+ SUBSTRING(@sValue, @iIndex2, 2) + '/' 
						+ SUBSTRING(@sValue, @iIndex3, 4);
				END

				IF (@iOperatorID = 1) 
				BEGIN
					/* Equal To. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' = ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL';
					END
				END

				IF (@iOperatorID = 2)
				BEGIN
					/* Not Equal To. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' <> ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL';
					END
				END

				IF (@iOperatorID = 3) 
				BEGIN
					/* Less than or Equal To. */
					IF LEN(@sValue) > 0 
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' <= ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL';
					END
				END

				IF (@iOperatorID = 4) 
				BEGIN
					/* Greater than or Equal To. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' >= ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL OR ' + @sColumnName + ' IS NOT NULL';
					END
				END
				
				IF (@iOperatorID = 5) 
				BEGIN
					/* Greater than. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' > ''' + @sValue + '''';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL';
					END
				END
				
				IF (@iOperatorID = 6)
				BEGIN
					/* Less than. */
					IF LEN(@sValue) > 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' < ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL';
					END
					ELSE
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL AND ' + @sColumnName + ' IS NOT NULL';
					END
				END
			END
			
			IF ((@iDataType <> -7) AND (@iDataType <> 2) AND (@iDataType <> 4) AND (@iDataType <> 11)) 
			BEGIN
				/* Character/Working Pattern column. */
				IF (@iOperatorID = 1) 
				BEGIN
					/* Equal To. */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' = '''' OR ' + @sColumnName + ' IS NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sValue = replace(@sValue, '*', '%');
						SET @sValue = replace(@sValue, '?', '_');
						SET @sSubFilterSQL = @sColumnName + ' LIKE ''' + @sValue + '''';
					END
				END
				
				IF (@iOperatorID = 2) 
				BEGIN
					/* Not Equal To. */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' <> '''' AND ' + @sColumnName + ' IS NOT NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sValue = replace(@sValue, '*', '%');
						SET @sValue = replace(@sValue, '?', '_');
						SET @sSubFilterSQL = @sColumnName + ' NOT LIKE ''' + @sValue + '''';
					END
				END

				IF (@iOperatorID = 7)
				BEGIN
					/* Contains */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL OR ' + @sColumnName + ' IS NOT NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sSubFilterSQL = @sColumnName + ' LIKE ''%' + @sValue + '%''';
					END
				END
				
				IF (@iOperatorID = 8) 
				BEGIN
					/* Does Not Contain. */
					IF LEN(@sValue) = 0
					BEGIN
						SET @sSubFilterSQL = @sColumnName + ' IS NULL AND ' + @sColumnName + ' IS NOT NULL';
					END
					ELSE
					BEGIN
						/* Replace the standard * and ? characters with the SQL % and _ characters. */
						SET @sValue = replace(@sValue, '''', '''''');
						SET @sSubFilterSQL = @sColumnName + ' NOT LIKE ''%' + @sValue + '%''';
					END
				END
			END
			
			IF LEN(@sSubFilterSQL) > 0
			BEGIN
				/* Add the filter code for this grid record into the complete filter code. */
				IF LEN(@sFilterSQL) > 0
				BEGIN
					SET @sFilterSQL = @sFilterSQL + ' AND (';
				END
				ELSE
				BEGIN
					SET @sFilterSQL = @sFilterSQL + '(';
				END
				SET @sFilterSQL = @sFilterSQL + @sSubFilterSQL + ')';
			END
		END
	END
	/* Loop through the tables used in the order, getting the column permissions for each one. */
	DECLARE tablesCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT DISTINCT ASRSysColumns.tableID
		FROM [dbo].[ASRSysOrderItems]
		INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
		WHERE ASRSysOrderItems.orderID = @piOrderID;

	OPEN tablesCursor;
	FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iTempTableID = @piTableID
		BEGIN
			/* Base table - use the real source. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				@sRealSource,
				syscolumns.name,
				MAX(CASE WHEN [protectType] IN (204, 205) AND [action] = 193 THEN 1	ELSE 0 END) AS selectGranted,
				MAX(CASE WHEN [protectType] IN (204, 205) AND [action] = 197 THEN 1	ELSE 0 END) AS updateGranted
			FROM @sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.action IN (193, 197)
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name = @sRealSource
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
			GROUP BY syscolumns.name,[protectType];
		END
		ELSE
		BEGIN
			/* Parent of the base table - get permissions for the table, and any associated views. */
			INSERT INTO @ColumnPermissions
			SELECT 
				@iTempTableID,
				sysobjects.name,
				syscolumns.name,
				MAX(CASE WHEN [protectType] IN (204, 205) AND [action] = 193 THEN 1	ELSE 0 END) AS selectGranted,
				MAX(CASE WHEN [protectType] IN (204, 205) AND [action] = 197 THEN 1	ELSE 0 END) AS updateGranted
			FROM @sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.action IN (193, 197)
				AND syscolumns.name <> 'timestamp'
				AND sysobjects.name IN (SELECT ASRSysTables.tableName FROM ASRSysTables WHERE 
					ASRSysTables.tableID = @iTempTableID 
					UNION SELECT ASRSysViews.viewName FROM ASRSysViews WHERE ASRSysViews.viewTableID = @iTempTableID)
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))
			GROUP BY sysobjects.name, syscolumns.name,[protectType];

		END
		FETCH NEXT FROM tablesCursor INTO @iTempTableID;
	END
	
	CLOSE tablesCursor;
	DEALLOCATE tablesCursor;

	/* Create the order select strings. */
	INSERT @FindDefinition
		SELECT c.tableID,
			o.columnID, 
			c.columnName,
	    	t.tableName,
			o.ascending,
			o.type,
			c.dataType,
			c.controltype,
			c.size,
			c.decimals,
			c.Use1000Separator,
			c.BlankIfZero,
			CASE WHEN ISNULL(o.Editable, 0) = 0 OR ISNULL(c.readonly, 1) = 1 THEN 0 ELSE 1 END,
			ISNULL(c.LookupTableID, 0) AS LookupTableID,
			ISNULL(c.LookupColumnID, 0) AS LookupColumnID,
			ISNULL(c.LookupFilterColumnID, 0) AS LookupFilterColumnID,
			ISNULL(c.LookupFilterValueID, 0) AS LookupFilterValueID,
			c.SpinnerMinimum,
			c.SpinnerMaximum,
			c.SpinnerIncrement,
			c.defaultvalue AS DefaultValue,
			c.Mask,
			c.dfltvalueexprid as DefaultValueExprID
		FROM [dbo].[ASRSysOrderItems] o
		INNER JOIN ASRSysColumns c ON o.columnID = c.columnID
		INNER JOIN ASRSysTables t ON c.tableID = t.tableID
		WHERE o.orderID = @piOrderID
		ORDER BY o.sequence;

	
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableID, columnID, columnName, tableName, ascending, type, datatype, size, decimals, Use1000Separator, BlankIfZero FROM @FindDefinition;


	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iSize, @iDecimals, @bUse1000Separator, @bBlankIfZero;

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
			FROM @ColumnPermissions
			WHERE tableViewName = @sRealSource
				AND columnName = @sColumnName;
				
			IF @fSelectGranted IS NULL SET @fSelectGranted = 0;
			IF @fSelectGranted = 1
			BEGIN
				/* The user DOES have SELECT permission on the column in the current table/view. */
				IF @sType = 'F'
				BEGIN
					/* Find column. */
					IF @iDataType = -4 OR @iDataType = -3
					BEGIN
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + 
						'RTRIM(SUBSTRING(CONVERT(VARCHAR(max),' + @sRealSource + '.' + @sColumnName + '), 11, 69)) AS [' + @sColumnName + ']';
					END
					ELSE
					BEGIN
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + 
						@sRealSource + '.' + @sColumnName;

					END

					SET @sSelectSQL = @sSelectSQL + @sTempString;
					INSERT INTO @OriginalColumns VALUES(@iColumnID, @sColumnName)
					SET @sThousandColumns = @sThousandColumns + convert(varchar(1),@bUse1000Separator);
					SET @sBlankIfZeroColumns = @sBlankIfZeroColumns + convert(varchar(1),@bBlankIfZero);
					
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType;
						SET @piColumnSize = @iSize;
						SET @piColumnDecimals = @iDecimals;
						SET @fFirstColumnAsc = @fAscending;
						SET @sFirstColCode = @sRealSource + '.' + @sColumnName;
						IF (@psAction = 'LOCATEID')
						BEGIN
							IF @piColumnType = 11 /* Date column */
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = isnull(convert(varchar(MAX), ' + @sRealSource + '.' + @sColumnName + ', 101), '''')';
							END
							ELSE
							IF @piColumnType = -7 /* Logic column */
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = convert(varchar(MAX), isnull(' + @sRealSource + '.' + @sColumnName + ', 0))';
							END
							ELSE IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = convert(varchar(MAX), isnull(' + @sRealSource + '.' + @sColumnName + ', 0))';
							END
							ELSE
							BEGIN
								SET @sTempExecString = 'SELECT @sTempLocateValue = replace(isnull(' + @sRealSource + '.' + @sColumnName + ', ''''), '''''''', '''''''''''')' ;
							END
							
							SET @sTempExecString = @sTempExecString
								+ ' FROM ' + @sRealSource
								+ ' WHERE ' + @sRealSource + '.ID = ' + @psLocateValue;
							SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
							EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
						END
					END
					
					SET @sOrderSQL = @sOrderSQL + CASE 
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
			FROM @ColumnPermissions
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
					IF @iDataType = -4 OR @iDataType = -3
					BEGIN
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + 
						'RTRIM(SUBSTRING(CONVERT(VARCHAR(max),' + @sRealSource + '.' + @sColumnName + '), 11, 69)) AS [' + @sColumnName + ']';						
					END
					ELSE
					BEGIN
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + @sColumnTableName + '.' + @sColumnName;
					END

					SET @sSelectSQL = @sSelectSQL + @sTempString;
					INSERT INTO @OriginalColumns VALUES(@iColumnID, @sColumnName)
					
				END
				ELSE
				BEGIN
					/* Order column. */
					IF len(@sOrderSQL) = 0 
					BEGIN
						SET @piColumnType = @iDataType;
						SET @piColumnSize = @iSize;
						SET @piColumnDecimals = @iDecimals;
						SET @fFirstColumnAsc = @fAscending;
						SET @sFirstColCode = @sColumnTableName + '.' + @sColumnName;
						IF (@psAction = 'LOCATEID')
						BEGIN
							SET @sTempLocateValue = '';
						END
					END
					SET @sOrderSQL = @sOrderSQL + CASE 
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
				FROM @JoinParents
				WHERE tableViewName = @sColumnTableName;
				
				IF @iTempCount = 0
				BEGIN
					INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sColumnTableName, @iColumnTableID);
				END
			END
			ELSE	
			BEGIN
				SET @sSubString = '';
				
				DECLARE viewCursor CURSOR LOCAL FAST_FORWARD FOR 
					SELECT tableViewName
					FROM @ColumnPermissions
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
					FROM @JoinParents
					WHERE tableViewname = @sViewName;
					
					IF @iTempCount = 0
					BEGIN
						INSERT INTO @JoinParents (tableViewName, tableID) VALUES(@sViewName, @iColumnTableId);
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
					IF @iDataType = -4 OR @iDataType = -3
					BEGIN
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
						END + 
						'RTRIM(SUBSTRING(CONVERT(VARCHAR(max),' + @sRealSource + '.' + @sColumnName + '), 11, 69)) AS [' + @sColumnName + ']';						
					END
					ELSE
					BEGIN
					SET @sTempString = CASE 
							WHEN (len(@sSelectSQL) > 0) THEN ',' 
							ELSE '' 
							END + 
							CASE
								WHEN @iDataType = 11 THEN 'convert(datetime, ' + @sSubString + ')'
								ELSE @sSubString 
							END;
					END
					
					SET @sSelectSQL = @sSelectSQL + @sTempString;
						
					SET @sTempString = ' AS [' + @sColumnName + ']';
					SET @sSelectSQL = @sSelectSQL + @sTempString;
					INSERT INTO @OriginalColumns VALUES(@iColumnID, @sColumnName)
						
					END
					ELSE
					BEGIN
						/* Order column. */
						IF len(@sOrderSQL) = 0 
						BEGIN
							SET @piColumnType = @iDataType;
							SET @piColumnSize = @iSize;
							SET @piColumnDecimals = @iDecimals;
							SET @fFirstColumnAsc = @fAscending;
							SET @sFirstColCode = @sSubString;
							IF (@psAction = 'LOCATEID')
							BEGIN
								SET @sTempLocateValue = '';
							END
						END
						SET @sOrderSQL = @sOrderSQL + CASE 
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
		FETCH NEXT FROM orderCursor INTO @iColumnTableId, @iColumnId, @sColumnName, @sColumnTableName, @fAscending, @sType, @iDataType, @iSize, @iDecimals, @bUse1000Separator, @bBlankIfZero;
	END

	CLOSE orderCursor;
	DEALLOCATE orderCursor;

	IF @psAction = 'LOCATEID'
	BEGIN
		SET @psAction = 'LOCATE';
		SET @psLocateValue = @sTempLocateValue;
	END
	
	/* Set the flags that show if no order columns could be selected, or if only some of them could be selected. */
	SET @pfSomeSelectable = CASE WHEN (len(@sSelectSQL) > 0) THEN 1 ELSE 0 END;
	SET @pfSomeNotSelectable = @fSelectDenied;

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
	SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource;
	IF @piParentTableID > 0 
	BEGIN
		SET @sTempExecString = @sTempExecString + 
			' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
		IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' AND ' + @sFilterSQL;
	END
	ELSE
	BEGIN
		IF len(@sFilterSQL) > 0	SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL;
	END
	
	SET @sTempParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @iCount OUTPUT;

	SET @piTotalRecCount = @iCount;
	IF len(@sSelectSQL) > 0 
	BEGIN
		SET @sTempString = ',' + @sRealSource + '.ID, CONVERT(int, ' + @sRealSource + '.Timestamp) AS [Timestamp]';
		SET @sSelectSQL = @sSelectSQL + @sTempString;
		INSERT INTO @OriginalColumns VALUES(0, 'ID') -- We don't need the ID
		INSERT INTO @OriginalColumns VALUES(0, 'Timestamp') -- We don't need the ID

		SET @sExecString = 'SELECT ' 
		IF @psAction = 'MOVEFIRST' OR @psAction = 'LOCATE' 
		BEGIN
			SET @sTempString = 'TOP ' + convert(varchar(100), @piRecordsRequired) + ' ';
			SET @sExecString = @sExecString + @sTempString;
		END
	
		SET @sTempString = @sSelectSQL + ' FROM ' + @sRealSource;
		SET @sExecString = @sExecString + @sTempString;

		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @JoinParents;
		
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
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource;
			SET @sExecString = @sExecString + @sTempString;
		END
		IF (@psAction = 'MOVENEXT') 
		BEGIN
			IF (@piFirstRecPos +  @piCurrentRecCount + @piRecordsRequired - 1) > @piTotalRecCount
			BEGIN
				SET @iGetCount = @piTotalRecCount - (@piCurrentRecCount + @piFirstRecPos - 1);
			END
			ELSE
			BEGIN
				SET @iGetCount = @piRecordsRequired;
			END
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource;
			SET @sExecString = @sExecString + @sTempString;
				
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos + @piCurrentRecCount + @piRecordsRequired - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource;
			SET @sExecString = @sExecString + @sTempString;
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
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @iGetCount) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
			SET @sExecString = @sExecString + @sTempString;
				
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID IN (SELECT TOP ' + convert(varchar(100), @piFirstRecPos - 1) + ' ' + @sRealSource + '.ID' +
				' FROM ' + @sRealSource
			SET @sExecString = @sExecString + @sTempString;

		END
		/* Add the filter code. */
		IF @piParentTableID > 0 
		BEGIN
			SET @sTempString = ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
			SET @sExecString = @sExecString + @sTempString;
				
			IF len(@sFilterSQL) > 0
			BEGIN
				SET @sTempString = ' AND ' + @sFilterSQL;
				SET @sExecString = @sExecString + @sTempString;
			END
		END
		ELSE
		BEGIN
			IF len(@sFilterSQL) > 0
			BEGIN
				SET @sTempString = ' WHERE ' + @sFilterSQL;
				SET @sExecString = @sExecString + @sTempString;
			END
		END
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
			SET @sLocateCode = '';
			IF (@piParentTableID > 0) OR (len(@sFilterSQL) > 0) 
			BEGIN
				SET @sLocateCode = @sLocateCode + ' AND (' + @sFirstColCode;
			END
			ELSE
			BEGIN
				SET @sLocateCode = @sLocateCode + ' WHERE (' + @sFirstColCode;
			END;
			
			SET @sJoinSQL = '';
			DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
			SELECT tableViewName, 
				tableID
			FROM @JoinParents;
			OPEN joinCursor;
			FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
			WHILE (@@fetch_status = 0)
			BEGIN
				SET @sJoinSQL = @sJoinSQL + 
					' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';
				FETCH NEXT FROM joinCursor INTO @sTableViewName, @iJoinTableID;
			END;
			
			CLOSE joinCursor;
			DEALLOCATE joinCursor;

			IF (@piColumnType = 12) OR (@piColumnType = -1) /* Character or Working Pattern column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = replace(isnull(' + @sFirstColCode + ', ''''), '''''''', '''''''''''')' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL						
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;

				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ''' + @psLocateValue + '''';
					IF len(@psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL';
					END;
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ''' + @sTempLocateValue + '''';
					END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ''' + @psLocateValue + ''' OR ' + 
						@sFirstColCode + ' LIKE ''' + @psLocateValue + '%'' OR ' + @sFirstColCode + ' IS NULL';
						
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' <= ''' + @sTempLocateValue + '''';
					END;
				END
			END
			IF @piColumnType = 11 /* Date column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = isnull(convert(varchar(MAX),' + @sFirstColCode + ', 101), '''')' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL						
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;

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
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						IF len(@sTempLocateValue) > 0
						BEGIN
							SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ''' + @sTempLocateValue + '''';
						END
					END;
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
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						IF len(@sTempLocateValue) > 0
						BEGIN
							SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' < ''' + @sTempLocateValue + '''';
						END
					END;
				END
			END
			IF @piColumnType = -7 /* Logic column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = convert(varchar(MAX), isnull(' + @sFirstColCode + ', 0))' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL	
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL	;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;
				
				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END;
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ' + 
							CASE
								WHEN @sTempLocateValue = 'True' THEN '1'
								ELSE '0'
							END;
					END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + 
						CASE
							WHEN @psLocateValue = 'True' THEN '1'
							ELSE '0'
						END;
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' < ' + 
							CASE
								WHEN @sTempLocateValue = 'True' THEN '1'
								ELSE '0'
							END;
					END;
				END
			END
			IF (@piColumnType = 2) OR (@piColumnType = 4) /* Numeric or Integer column */
			BEGIN
				IF @psOriginalAction = 'LOCATEID'	
				BEGIN
					SET @sTempExecString = 'SELECT TOP 1 @sTempLocateValue = convert(varchar(MAX), isnull(' + @sFirstColCode + ', 0))' 
						+ ' FROM ' + @sRealSource
						+ @sJoinSQL	
						+ ' WHERE ' + @sRealSource + '.ID IN (' 
						+ '     SELECT TOP ' + convert(varchar(100), @piRecordsRequired) + ' ' + @sRealSource + '.ID'
						+ '     FROM ' + @sRealSource
						+ @sJoinSQL;

					SET @sTempFilterString = '';
					IF @piParentTableID > 0 SET @sTempExecString = @sTempExecString + ' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);						
					IF LEN(@sFilterSQL) > 0
					BEGIN
						IF LEN(@sTempFilterString) > 0
							SET @sTempFilterString = @sTempFilterString + ' AND ' + @sFilterSQL	;
						ELSE
							SET @sTempFilterString = @sTempFilterString + ' WHERE ' + @sFilterSQL	;

						SET @sTempExecString = @sTempExecString + @sTempFilterString;
					END;
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sReverseOrderSQL + ')';
					SET @sTempExecString = @sTempExecString +  ' ORDER BY ' + @sOrderSQL ;

					SET @sTempParamDefinition = N'@sTempLocateValue varchar(MAX) OUTPUT';
					EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @sTempLocateValue OUTPUT;
				END;

				IF @fFirstColumnAsc = 1
				BEGIN
					SET @sLocateCode = @sLocateCode + ' >= ' + @psLocateValue;
					IF convert(float, @psLocateValue) = 0
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' IS NULL';
					END
					
					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' > ' + @sTempLocateValue;
					END;
				END
				ELSE
				BEGIN
					SET @sLocateCode = @sLocateCode + ' <= ' + @psLocateValue + ' OR ' + @sFirstColCode + ' IS NULL';

					IF @psOriginalAction = 'LOCATEID'	
					BEGIN
						SET @sLocateCode = @sLocateCode + ' OR ' + @sFirstColCode + ' < ' + @sTempLocateValue ;
					END;
				END
			END
			SET @sLocateCode = @sLocateCode + ')';
			SET @sTempString = @sLocateCode;
			SET @sExecString = @sExecString + @sTempString;
		END

		/* Add the ORDER BY code to the find record selection string if required. */
		IF (@RecordID = -1) BEGIN --Return all records
			SET @sTempString = ' ORDER BY ' + @sOrderSQL;
		END ELSE BEGIN --Return only the requested record
			IF charindex(' WHERE ', @sExecString) > 0 BEGIN --A WHERE clause may have been added for the filtering of records
				SET @sTempString = ' AND '
			END ELSE BEGIN
				SET @sTempString = ' WHERE '
			END

			SET @sTempString = @sTempString + @sRealSource + '.ID = ' + CONVERT(varchar(100), @RecordID) + ' ORDER BY ' + @sOrderSQL;
		END

		SET @sExecString = @sExecString + @sTempString;
				
	END

	/* Check if the user has insert or delete permission on the table. */
	SET @pfInsertGranted = 0;
	SET @pfDeleteGranted = 0;
	IF LEN(@sRealSource) > 0
	BEGIN
		DECLARE tableInfo_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT p.action
		FROM @sysprotects p
		INNER JOIN sysobjects ON p.id = sysobjects.id
		WHERE p.protectType <> 206
			AND ((p.action = 195) OR (p.action = 196))
			AND sysobjects.name = @sRealSource;
			
		OPEN tableInfo_cursor;
		FETCH NEXT FROM tableInfo_cursor INTO @iTempAction;
		WHILE (@@fetch_status = 0)
		BEGIN
			IF @iTempAction = 195
			BEGIN
				SET @pfInsertGranted = 1;
			END
			ELSE
			BEGIN
				SET @pfDeleteGranted = 1;
			END
			FETCH NEXT FROM tableInfo_cursor INTO @iTempAction;
		END
		
		CLOSE tableInfo_cursor;
		DEALLOCATE tableInfo_cursor;
		
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
		SET @sTempExecString = 'SELECT @recordCount = COUNT(' + @sRealSource + '.id) FROM ' + @sRealSource;
		DECLARE joinCursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT tableViewName, 
			tableID
		FROM @JoinParents;

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
		
		IF @piParentTableID > 0 
		BEGIN
			SET @sTempExecString = @sTempExecString + 
				' WHERE ' + @sRealSource + '.ID_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
			IF len(@sFilterSQL) > 0 SET @sTempExecString = @sTempExecString + ' AND ' + @sFilterSQL;
		END
		ELSE
		BEGIN
			IF len(@sFilterSQL) > 0	SET @sTempExecString = @sTempExecString + ' WHERE ' + @sFilterSQL;
		END
		
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
				WHEN @piTotalRecCount < @piFirstRecPos + @piRecordsRequired THEN 1
				ELSE 0
			END;
	END
	
	-- Return a recordset of the required columns in the required order from the given table/view.
	IF (@pfSomeSelectable = 1)
	BEGIN
		IF @piViewID <>0 BEGIN
			SELECT @sTableOrViewName = ViewName FROM ASRSysViews WHERE ViewID = @piViewID
		END ELSE IF @piTableID <> 0 BEGIN
			SELECT @sTableOrViewName = TableName FROM ASRSysTables WHERE TableID = @piTableID
		END

		SELECT @sBlankIfZeroColumns AS BlankIfZeroColumns
			, @sThousandColumns AS ThousandColumns, @sTableOrViewName AS TableOrViewName

		EXECUTE sp_executeSQL @sExecString;

		SELECT f.tableID, f.columnID, f.columnName, f.ascending, f.type, f.datatype, f.controltype, f.size, f.decimals, f.Use1000Separator, f.BlankIfZero
			 , CASE WHEN f.Editable = 1 AND p.updateGranted = 1 THEN 1 ELSE 0 END AS updateGranted
			 , LookupTableID, LookupColumnID, LookupFilterColumnID, LookupFilterValueID
			 ,SpinnerMinimum, SpinnerMaximum, SpinnerIncrement, DefaultValue, Mask, DefaultValueExprID
			FROM @FindDefinition f
				INNER JOIN @ColumnPermissions p ON p.columnName = f.columnName
			WHERE f.[type] = 'F';

		SELECT columnID, columnName FROM @OriginalColumns ORDER BY columnName
	END

END
