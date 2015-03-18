CREATE PROCEDURE [dbo].[spASRIntGetRecord] (
	@piRecordID			integer 		OUTPUT,
	@piRecordCount		integer 		OUTPUT,
	@piRecordPosition	integer 		OUTPUT,
	@psFilterDef 		varchar(MAX),
	@psAction	 		varchar(100),
	@piParentTableID	integer,
	@piParentRecordID	integer,
	@psDecimalSeparator	varchar(100),
	@psLocaleDateFormat	varchar(100),
	@piScreenID 		integer,
	@piViewID 			integer,
	@piOrderID			integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iRecordID 				integer, 
		@iRecordCount 				integer,
		@iRecordPosition 			integer,
		@sCommand					nvarchar(MAX),
		@sLongCommand				nvarchar(MAX),
		@sParamDefinition			nvarchar(500),
		@sSubCommand				nvarchar(MAX),
		@sSubParamDefinition		nvarchar(500),
		@sPositionCommand			nvarchar(MAX),
		@sTemp						nvarchar(MAX),
		@sPositionParamDefinition	nvarchar(500),
		@sMoveCommand				nvarchar(MAX),
		@sReverseOrderSQL			varchar(MAX),
		@sRelevantOrderSQL			varchar(MAX),
		@sRemainingSQL				varchar(MAX),
		@iCharIndex					integer,
		@iLastCharIndex				integer,
		@sDESCstring				varchar(5),
		@fPositionKnown				bit,
		@sPreviousWhere				varchar(MAX),
		@sOrderItem					varchar(MAX),
		@sOrderColumn				varchar(MAX),
		@sOrderTable				varchar(MAX),
		@iDotIndex 					integer,
		@iDataType					integer,
		@fBitValue					bit,
		@sVarCharValue				varchar(MAX),
		@iIntValue					integer,
		@dblNumValue				float,
		@dtDateValue				datetime,
		@sTempTableName				sysname,
		@sTempTablePrefix			sysname,
		@iSpaceIndex 				integer,
		@fDescending				integer,
		@fAddedToPositionString		bit,
		@fAddedToMoveString			bit,
		@sSubString 				varchar(MAX),
		@sTempName 					sysname,
		@sPositionSQL				nvarchar(MAX),
		@sFromSQL					varchar(MAX),
		@sRealSource				varchar(MAX),
		@iIndex1					integer,
		@iIndex2					integer,
		@iIndex3					integer,
		@iColumnID					integer,
		@iOperatorID				integer,
		@sValue						varchar(MAX),
		@sFilterSQL					nvarchar(MAX),
		@sSubFilterSQL				nvarchar(MAX),
		@sColumnName 				sysname,
		@sTableViewName				sysname,
		@iJoinTableID				integer,
		@sFromDef					varchar(MAX),
		@sSelectSQL					nvarchar(MAX),
		@sExecuteSQL				nvarchar(MAX),
		@sOrderSQL 					varchar(MAX);

	exec [dbo].[spASRIntGetScreenStrings]
		@piScreenID,
		@piViewID,
		@sSelectSQL output,
		@sFromDef output,
		@sOrderSQL output,
		@piOrderID output;

	SET @sFilterSQL = '';
	SET @sPositionCommand = '';
	SET @fPositionKnown = 0;
	SET @sDESCstring = ' DESC';
	SET @iRecordID = @piRecordID;
	SET @fAddedToPositionString = 0;
	SET @fAddedToMoveString = 0;
	SET @iIndex1 = charindex('	', @sFromDef);
	SET @sRealSource = replace(LEFT(@sFromDef, @iIndex1-1), '''', '''''');
	SET @sFromSQL = @sRealSource;
	SET @sFromDef = SUBSTRING(@sFromDef, @iIndex1+1, LEN(@sFromDef) - @iIndex1);

	WHILE charindex('	', @sFromDef) > 0
	BEGIN
		SET @iIndex1 = charindex('	', @sFromDef);
		SET @iIndex2 = charindex('	', @sFromDef, @iIndex1+1);
				
		SET @sTableViewName = replace(LEFT(@sFromDef, @iIndex1-1), '''', '''''');
		SET @iJoinTableID = convert(integer, SUBSTRING(@sFromDef, @iIndex1+1, @iIndex2-@iIndex1-1));
				
		SET @sFromDef = SUBSTRING(@sFromDef, @iIndex2+1, LEN(@sFromDef) - @iIndex2);
		SET @sFromSQL = @sFromSQL + 
			' LEFT OUTER JOIN ' + convert(varchar(255), @sTableViewName) + ' ON ' + @sRealSource + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + convert(varchar(255), @sTableViewName) + '.ID';

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
			FROM ASRSysColumns
			WHERE columnID = @iColumnID;
							
			SET @sColumnName = @sRealSource + '.' + @sColumnName;

			IF (@iDataType = -7) 
			BEGIN
				/* Logic column (must be the equals operator).	*/
				SET @sSubFilterSQL = @sColumnName + ' = ';
			
				IF UPPER(@sValue) = 'TRUE'
					SET @sSubFilterSQL = @sSubFilterSQL + '1';
				ELSE
					SET @sSubFilterSQL = @sSubFilterSQL + '0';
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
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
				END

				IF (@iOperatorID = 2)
				BEGIN
					/* Not Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' <> ' + @sValue;
					IF convert(float, @sValue) = 0
						SET @sSubFilterSQL = @sSubFilterSQL + ' AND ' + @sColumnName + ' IS NOT NULL';
				END

				IF (@iOperatorID = 3) 
				BEGIN
					/* Less than or Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' <= ' + @sValue;
					IF convert(float, @sValue) = 0
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
				END
        
				IF (@iOperatorID = 4) 
				BEGIN
					/* Greater than or Equal To. */
					SET @sSubFilterSQL = @sColumnName + ' >= ' + @sValue;
					IF convert(float, @sValue) = 0
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
				END
				
				IF (@iOperatorID = 5) 
				BEGIN
					/* Greater than. */
					SET @sSubFilterSQL = @sColumnName + ' > ' + @sValue;
					IF convert(float, @sValue) = 0
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
				END
				
				IF (@iOperatorID = 6) 
				BEGIN
					/* Less than.*/
					SET @sSubFilterSQL = @sColumnName + ' < ' + @sValue;
					IF convert(float, @sValue) = 0
						SET @sSubFilterSQL = @sSubFilterSQL + ' OR ' + @sColumnName + ' IS NULL';
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
						SET @sSubFilterSQL = @sColumnName + ' = ''' + @sValue + '''';
					ELSE
						SET @sSubFilterSQL = @sColumnName + ' IS NULL';
				END

				IF (@iOperatorID = 2)
				BEGIN
					/* Not Equal To. */
					IF LEN(@sValue) > 0
						SET @sSubFilterSQL = @sColumnName + ' <> ''' + @sValue + ''''
					ELSE
						SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL';
				END

				IF (@iOperatorID = 3) 
				BEGIN
					/* Less than or Equal To. */
					IF LEN(@sValue) > 0 
						SET @sSubFilterSQL = @sColumnName + ' <= ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL';
					ELSE
						SET @sSubFilterSQL = @sColumnName + ' IS NULL';
				END

				IF (@iOperatorID = 4) 
				BEGIN
					/* Greater than or Equal To. */
					IF LEN(@sValue) > 0
						SET @sSubFilterSQL = @sColumnName + ' >= ''' + @sValue + ''''
					ELSE
						SET @sSubFilterSQL = @sColumnName + ' IS NULL OR ' + @sColumnName + ' IS NOT NULL';
				END

				IF (@iOperatorID = 5) 
				BEGIN
					/* Greater than. */
					IF LEN(@sValue) > 0
						SET @sSubFilterSQL = @sColumnName + ' > ''' + @sValue + '''';
					ELSE
						SET @sSubFilterSQL = @sColumnName + ' IS NOT NULL';
				END

				IF (@iOperatorID = 6)
				BEGIN
					/* Less than. */
					IF LEN(@sValue) > 0
						SET @sSubFilterSQL = @sColumnName + ' < ''' + @sValue + ''' OR ' + @sColumnName + ' IS NULL';
					ELSE
						SET @sSubFilterSQL = @sColumnName + ' IS NULL AND ' + @sColumnName + ' IS NOT NULL';
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
					SET @sFilterSQL = @sFilterSQL + ' AND (';
				ELSE
					SET @sFilterSQL = @sFilterSQL + '(';
					
				SET @sFilterSQL = @sFilterSQL + @sSubFilterSQL + ')';
			END
		END
	END

	IF (@psAction = 'LOAD') AND (@piRecordID = 0) SET @psAction = 'MOVEFIRST';

	IF (@psAction = 'LOAD') AND (@piRecordID > 0) 
	BEGIN
		/* Check the required record is still in the recordset. */
		SET @sSubCommand = 'SELECT @iValue = ' + @sRealSource + '.ID ' + 
			' FROM ' + @sRealSource +
			' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @piRecordID);
		IF len(@sFilterSQL) > 0
		BEGIN
			SET @sSubCommand = @sSubCommand + 
				' AND ' + @sFilterSQL;
		END
		SET @sSubParamDefinition = N'@iValue integer OUTPUT';
		EXEC sp_executesql @sSubCommand,  @sSubParamDefinition, @iIntValue OUTPUT;
		
		IF @iIntValue IS NULL 
			SET @psAction = 'MOVEFIRST';

	END

	/* Create the reverse order SQL if required. */
	SET @sReverseOrderSQL = '';
	IF (@psAction = 'MOVELAST') OR (@psAction = 'MOVEPREVIOUS')
	BEGIN
		SET @sRemainingSQL = @sOrderSQL;
		SET @iLastCharIndex = 0;
		SET @iCharIndex = CHARINDEX(', ', @sOrderSQL);
		WHILE @iCharIndex > 0 
		BEGIN
			IF UPPER(SUBSTRING(@sOrderSQL, @iCharIndex - LEN(@sDESCstring), LEN(@sDESCstring))) = @sDESCstring
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - LEN(@sDESCstring) - @iLastCharIndex) + ', ';
			ELSE
				SET @sReverseOrderSQL = @sReverseOrderSQL + SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - @iLastCharIndex) + @sDESCstring + ', ';

			SET @iLastCharIndex = @iCharIndex;
			SET @iCharIndex = CHARINDEX(', ', @sOrderSQL, @iLastCharIndex + 1);
	
			SET @sRemainingSQL = SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, LEN(@sOrderSQL) - @iLastCharIndex);
		END
		SET @sReverseOrderSQL = @sReverseOrderSQL + @sRemainingSQL + @sDESCstring;
	END

	/* Get the record count of the required recordset. */	
	SET @sCommand = 'SELECT @recordCount = COUNT(id)' +
		' FROM ' + @sRealSource;

	IF @piParentTableID > 0
	BEGIN
		SET @sCommand = @sCommand +
			' WHERE ' + @sRealSource + '.id_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
		IF len(@sFilterSQL) > 0
			SET @sCommand = @sCommand + ' AND ' + @sFilterSQL;
	END
	ELSE
	BEGIN
		IF len(@sFilterSQL) > 0
			SET @sCommand = @sCommand + ' WHERE ' + @sFilterSQL;
	END

	SET @sParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordCount OUTPUT;

	SET @piRecordCount = @iRecordCount;
	
	/* Get the required record ID and record position values if we're moving to the first or last records. */
	IF (@psAction = 'MOVEFIRST') OR (@psAction = 'MOVELAST')
	BEGIN
		SET @fPositionKnown = 1;
		SET @sLongCommand = 'SELECT TOP 1 @recordID = ' + @sRealSource + '.id' + ' FROM ' + @sFromSQL;
		IF @piParentTableID > 0
		BEGIN
			SET @sLongCommand = @sLongCommand +
				' WHERE ' + @sRealSource + '.id_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
			IF len(@sFilterSQL) > 0
				SET @sLongCommand = @sLongCommand +	' AND ' + @sFilterSQL;
		END
		ELSE
		BEGIN
			IF len(@sFilterSQL) > 0
				SET @sLongCommand = @sLongCommand +	' WHERE ' + @sFilterSQL;
		END

		SET @sLongCommand = @sLongCommand +
			' ORDER BY ' + 
			CASE 
				WHEN @psAction = 'MOVEFIRST' THEN @sOrderSQL
				ELSE @sReverseOrderSQL
			END;

		SET @sParamDefinition = N'@recordID integer OUTPUT';
		EXEC sp_executesql @sLongCommand,  @sParamDefinition, @iRecordID OUTPUT;

		IF @iRecordID IS NULL 
		BEGIN
			SET @iRecordID = 0;
		END
		SET @iRecordPosition = 
			CASE
				WHEN (@psAction = 'MOVEFIRST') AND (@iRecordCount > 0) THEN 1
				ELSE @iRecordCount
			END
	END
	
	/* Get the required record ID and record position values if we're moving to the next or previous records. */
	IF (@psAction = 'MOVENEXT') OR (@psAction = 'MOVEPREVIOUS')
	BEGIN

		SET @sMoveCommand = 'SELECT TOP 1 @recordID = ' + @sRealSource + '.id' +
			' FROM ' + @sFromSQL + 
			' WHERE ';

		IF @piParentTableID > 0
		BEGIN
			SET @sTemp = @sRealSource + '.id_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID) + ' AND ';
			SET @sMoveCommand = @sMoveCommand + @sTemp;
		END

		IF len(@sFilterSQL) > 0
		BEGIN
			SET @sTemp = @sFilterSQL + ' AND ';
			SET @sMoveCommand = @sMoveCommand + @sTemp;
		END		

		SET @sRelevantOrderSQL = CASE WHEN @psAction = 'MOVENEXT' THEN @sOrderSQL ELSE @sReverseOrderSQL END;
		SET @sPreviousWhere = '';
		SET @sTemp = 	'(';
		SET @sMoveCommand = @sMoveCommand + @sTemp;

		/* Get the order column values for the current record. */
		SET @iLastCharIndex = 0;
		SET @iCharIndex = CHARINDEX(', ', @sRelevantOrderSQL);

		WHILE @iCharIndex > 0 
		BEGIN
			SET @fDescending = 
				CASE
					WHEN UPPER(SUBSTRING(@sRelevantOrderSQL, @iCharIndex - LEN(@sDESCstring), len(@sDESCstring))) = @sDESCstring THEN 1
					ELSE 0
				END
			SET @sOrderItem = SUBSTRING(@sRelevantOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1- (@fDescending * LEN(@sDESCstring)) - @iLastCharIndex);
			SET @iDotIndex = CHARINDEX('.', @sOrderItem);
			SET @sOrderTable = LTRIM(LEFT(@sOrderItem, @iDotIndex - 1));
			SET @iSpaceIndex = CHARINDEX(' ', REVERSE(@sOrderTable));

			IF @iSpaceIndex > 0 
				SET @sOrderTable = SUBSTRING(@sOrderTable, LEN(@sOrderTable) - @iSpaceIndex + 2, @iSpaceIndex - 1);

			SET @sOrderColumn = RTRIM(SUBSTRING(@sOrderItem, @iDotIndex + 1, LEN(@sOrderItem) - @iDotIndex));
			SET @iSpaceIndex = CHARINDEX(' ', @sOrderColumn);

			IF @iSpaceIndex > 0 
				SET @sOrderColumn = SUBSTRING(@sOrderColumn, 1, @iSpaceIndex - 1);

			/* Get the data type of the order. */
			SELECT @iDataType = xtype
			FROM syscolumns
			WHERE name = @sOrderColumn
				AND id = (SELECT id FROM sysobjects WHERE name = @sOrderTable);

			IF @iDataType = 104	/* bit */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @fValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID);
				SET @sSubParamDefinition = N'@fValue bit OUTPUT';
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @fBitValue OUTPUT;

				IF @fBitValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = '(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = ' OR ('  + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)';
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0
							SET @sTemp = '(' + @sOrderItem + ' > ' + convert(varchar(2), @fBitValue) + ')';
						ELSE
							SET @sTemp = '((' + @sOrderItem + ' < ' + convert(varchar(2), @fBitValue) + ') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ' + convert(varchar(2), @fBitValue) + ')';
					END
					ELSE
					BEGIN
						IF @fDescending = 0
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ' + convert(varchar(2), @fBitValue) + '))';
						ELSE
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ' + convert(varchar(2), @fBitValue) + ') OR ('  + @sOrderItem + ' IS NULL)))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ' + convert(varchar(2), @fBitValue) + ')';
					END
				END
			END

			IF @iDataType = 167	/* varchar */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @sValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID);
				SET @sSubParamDefinition = N'@sValue varchar(MAX) OUTPUT';
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @sVarCharValue OUTPUT;

				IF @sVarCharValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = '(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)';
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
							SET @sTemp = '(' + @sOrderItem + ' > ''' + REPLACE(@sVarCharValue, '''', '''''')  + ''')';
						ELSE
							SET @sTemp = '((' + @sOrderItem + ' < ''' + REPLACE(@sVarCharValue, '''', '''''')  + ''') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ''' + REPLACE(@sVarCharValue, '''', '''''') + ''')';
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ''' +REPLACE(@sVarCharValue, '''', '''''') + '''))';
						ELSE
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ''' +REPLACE(@sVarCharValue, '''', '''''') + ''') OR ('  + @sOrderItem + ' IS NULL)))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ''' + REPLACE(@sVarCharValue, '''', '''''') + ''')';
					END
				END
			END

			IF @iDataType = 56	/* integer */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @iValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID);
				SET @sSubParamDefinition = N'@iValue integer OUTPUT';
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @iIntValue OUTPUT;

				IF @iIntValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = '(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)';
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
							SET @sTemp = '(' + @sOrderItem + ' > ' + convert(varchar(200), @iIntValue)  + ')';
						ELSE
							SET @sTemp = '((' + @sOrderItem + ' < ' + convert(varchar(200), @iIntValue)  + ') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ' + convert(varchar(200), @iIntValue) + ')'		;
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ' + convert(varchar(200), @iIntValue) + '))';
						ELSE
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ' + convert(varchar(200), @iIntValue) + ') OR ('  + @sOrderItem + ' IS NULL)))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ' + convert(varchar(200), @iIntValue) + ')';
					END
				END
			END

			IF @iDataType = 108	/* numeric */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @dblValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID);
				SET @sSubParamDefinition = N'@dblValue float OUTPUT';
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @dblNumValue OUTPUT;

				IF @dblNumValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = '(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)';
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
							SET @sTemp = '(' + @sOrderItem + ' > ' + convert(varchar(200), @dblNumValue)  + ')';
						ELSE
							SET @sTemp = '((' + @sOrderItem + ' < ' + convert(varchar(200), @dblNumValue)  + ') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ' + convert(varchar(200), @dblNumValue) + ')';
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ' + convert(varchar(200), @dblNumValue) + '))';
						ELSE
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ' + convert(varchar(200), @dblNumValue) + ') OR ('  + @sOrderItem + ' IS NULL)))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ' + convert(varchar(200), @dblNumValue) + ')';
					END
				END
			END

			IF @iDataType = 61	/* datetime */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @dtValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID);
				SET @sSubParamDefinition = N'@dtValue datetime OUTPUT';
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @dtDateValue OUTPUT;

				IF @dtDateValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = '(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
						BEGIN
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';
							SET @sMoveCommand = @sMoveCommand + @sTemp;
							SET @fAddedToMoveString = 1;
						END
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)';
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 0 
							SET @sTemp = '(' + @sOrderItem + ' > ''' + convert(varchar(50), @dtDateValue, 121)  + ''')';
						ELSE
							SET @sTemp = '((' + @sOrderItem + ' < ''' + convert(varchar(50), @dtDateValue, 121)  + ''') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ''' + convert(varchar(50), @dtDateValue, 121) + ''')';
					END
					ELSE
					BEGIN
						IF @fDescending = 0 
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ''' + convert(varchar(50), @dtDateValue, 121) + '''))';
						ELSE
							SET @sTemp = ' OR (' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ''' + convert(varchar(50), @dtDateValue, 121) + ''') OR ('  + @sOrderItem + ' IS NULL)))';

						SET @sMoveCommand = @sMoveCommand + @sTemp;
						SET @fAddedToMoveString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ''' + convert(varchar(50), @dtDateValue, 121) + ''')';
					END
				END
			END
	
			SET @iLastCharIndex = @iCharIndex;
			SET @iCharIndex = CHARINDEX(', ', @sRelevantOrderSQL, @iLastCharIndex + 1);
			SET @sRemainingSQL = SUBSTRING(@sRelevantOrderSQL, @iLastCharIndex + 2, len(@sRelevantOrderSQL) - @iLastCharIndex);
		END

		/* Add on the ID condition. */
		IF (@psAction = 'MOVENEXT')
		BEGIN
			IF LEN(@sPreviousWhere) = 0
			BEGIN
				SET @sTemp = '(' + @sRealSource + '.id > ' + convert(varchar(255), @iRecordID)  + ')';
			END
			ELSE
			BEGIN
				IF @fAddedToMoveString = 0
					SET @sTemp = ' (' + @sPreviousWhere + ' AND (' + @sRealSource + '.id > ' + convert(varchar(255), @iRecordID) + '))';
				ELSE
					SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (' + @sRealSource + '.id > ' + convert(varchar(255), @iRecordID) + '))';
			END
		END
		ELSE
		BEGIN
			IF LEN(@sPreviousWhere) = 0
			BEGIN
				SET @sTemp = '(' + @sRealSource + '.id < ' + convert(varchar(255), @iRecordID)  + ')';
			END
			ELSE
			BEGIN
				IF @fAddedToMoveString = 0
					SET @sTemp = ' (' + @sPreviousWhere + ' AND (' + @sRealSource + '.id < ' + convert(varchar(255), @iRecordID) + '))';
				ELSE
					SET @sTemp = ' OR (' + @sPreviousWhere + ' AND (' + @sRealSource + '.id < ' + convert(varchar(255), @iRecordID) + '))';
			END
		END

		SET @sTemp = @sTemp + ')';
		SET @sMoveCommand = @sMoveCommand + @sTemp;
		SET @fAddedToMoveString = 1;
		SET @sTemp = ' ORDER BY ' + @sRelevantOrderSQL;
		SET @sMoveCommand = @sMoveCommand + @sTemp;
		SET @fAddedToMoveString = 1;

		SET @sParamDefinition = N'@recordID integer OUTPUT';
		EXEC sp_executesql @sMoveCommand,  @sParamDefinition, @iRecordID OUTPUT;

		IF @iRecordID IS NULL 
			SET @iRecordID = 0;
			
	END

	IF @fPositionKnown = 0
	BEGIN
		/* Calculate the current record's position. */
		EXECUTE sp_ASRUniqueObjectName @sTempName OUTPUT, 'ASRSysTempInt', 3;
		EXECUTE ('CREATE TABLE ' + @sTempName + ' (result INT)');

		/* Calculate the current record's position. */
		SET @sPositionCommand = 'INSERT INTO ' + convert(varchar(255), @sTempName) + ' SELECT COUNT(' + @sRealSource + '.id)' +
			' FROM ' + @sFromSQL + 
			' WHERE ';

		IF @piParentTableID > 0
		BEGIN
			SET @sPositionCommand = @sPositionCommand +
				'(' + @sRealSource + '.id_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID) + ') AND '
		END

		IF len(@sFilterSQL) > 0
			SET @sPositionCommand = @sPositionCommand + '(' + @sFilterSQL + ') AND ';

		SET @sPositionCommand = @sPositionCommand + '(';
		SET @sPreviousWhere = '';

		/* Get the order column values for the current record. */
		SET @iLastCharIndex = 0;
		SET @iCharIndex = CHARINDEX(', ', @sOrderSQL);

		WHILE @iCharIndex > 0 
		BEGIN
			SET @fDescending = CASE
					WHEN UPPER(SUBSTRING(@sOrderSQL, @iCharIndex - LEN(@sDESCstring), LEN(@sDESCstring))) = @sDESCstring THEN 1
					ELSE 0
				END
			SET @sOrderItem = SUBSTRING(@sOrderSQL, @iLastCharIndex + 1, @iCharIndex - 1 - (@fDescending * LEN(@sDESCstring)) - @iLastCharIndex)
			SET @iDotIndex = CHARINDEX('.', @sOrderItem)
			SET @sOrderTable = LTRIM(LEFT(@sOrderItem, @iDotIndex - 1))
			SET @iSpaceIndex = CHARINDEX(' ', REVERSE(@sOrderTable))

			IF @iSpaceIndex > 0 
			BEGIN
				SET @sOrderTable = SUBSTRING(@sOrderTable, LEN(@sOrderTable) - @iSpaceIndex + 2, @iSpaceIndex - 1)
			END

			SET @sOrderColumn = RTRIM(SUBSTRING(@sOrderItem, @iDotIndex + 1, LEN(@sOrderItem) - @iDotIndex))
			SET @iSpaceIndex = CHARINDEX(' ', @sOrderColumn)

			IF @iSpaceIndex > 0 
			BEGIN
				SET @sOrderColumn = SUBSTRING(@sOrderColumn, 1, @iSpaceIndex - 1)
			END

			/* Get the data type of the order. */
			SELECT @iDataType = xtype
			FROM syscolumns
			WHERE name = @sOrderColumn
				AND id = (SELECT id FROM sysobjects WHERE name = @sOrderTable)

			IF @iDataType = 104	/* bit */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @fValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N'@fValue bit OUTPUT'
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @fBitValue OUTPUT

				IF @fBitValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = 	'(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END

						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'('  + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';

							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END

						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)';
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
							SET @sTemp = 	'(' + @sOrderItem + ' > ' + convert(varchar(MAX), @fBitValue) + ')';
						ELSE
							SET @sTemp = 	'((' + @sOrderItem + ' < ' + convert(varchar(MAX), @fBitValue) + ') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ' + convert(varchar(MAX), @fBitValue) + ')';
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sTemp = 	CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ' + convert(varchar(MAX), @fBitValue) + '))';
						END
						ELSE
						BEGIN
							SET @sTemp =CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ' + convert(varchar(MAX), @fBitValue) + ') OR ('  + @sOrderItem + ' IS NULL)))';
						END

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ' + convert(varchar(MAX), @fBitValue) + ')';
					END
				END
			END

			IF @iDataType = 167	/* varchar */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @sValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID)
				SET @sSubParamDefinition = N'@sValue varchar(MAX) OUTPUT'
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @sVarCharValue OUTPUT

				IF @sVarCharValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = 	'(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END

						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';

							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END

						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)'
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
							SET @sTemp = 	'(' + @sOrderItem + ' > ''' + REPLACE(@sVarCharValue, '''', '''''')  + ''')';
						ELSE
							SET @sTemp = 	'((' + @sOrderItem + ' < ''' + REPLACE(@sVarCharValue, '''', '''''')  + ''') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ''' + REPLACE(@sVarCharValue, '''', '''''') + ''')';
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sTemp = 	CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ''' +REPLACE(@sVarCharValue, '''', '''''') + '''))'
						END
						ELSE
						BEGIN
							SET @sTemp = 	CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ''' +REPLACE(@sVarCharValue, '''', '''''') + ''') OR ('  + @sOrderItem + ' IS NULL)))'
						END

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ''' + REPLACE(@sVarCharValue, '''', '''''') + ''')';

					END
				END
			END

			IF @iDataType = 56	/* integer */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @iValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID);
				SET @sSubParamDefinition = N'@iValue integer OUTPUT';
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @iIntValue OUTPUT;

				IF @iIntValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = 	'(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END
						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';

							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END

						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)';
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
							SET @sTemp = 	'(' + @sOrderItem + ' > ' + convert(varchar(MAX), @iIntValue)  + ')';
						ELSE
							SET @sTemp = 	'((' + @sOrderItem + ' < ' + convert(varchar(MAX), @iIntValue)  + ') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ' + convert(varchar(MAX), @iIntValue) + ')';
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sTemp = CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ' + convert(varchar(MAX), @iIntValue) + '))';
						END
						ELSE
						BEGIN
							SET @sTemp = CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ' + convert(varchar(MAX), @iIntValue) + ') OR ('  + @sOrderItem + ' IS NULL)))'
						END

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ' + convert(varchar(MAX), @iIntValue) + ')';
					END
				END
			END

			IF @iDataType = 108	/* numeric */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @dblValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID);
				SET @sSubParamDefinition = N'@dblValue float OUTPUT';
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @dblNumValue OUTPUT;

				IF @dblNumValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = '(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END

						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';

							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END

						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)';
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
							SET @sTemp = 	'(' + @sOrderItem + ' > ' + convert(varchar(200), @dblNumValue)  + ')';
						ELSE
							SET @sTemp = 	'((' + @sOrderItem + ' < ' + convert(varchar(200), @dblNumValue)  + ') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ' + convert(varchar(200), @dblNumValue) + ')';
						
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sTemp = CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ' + convert(varchar(200), @dblNumValue) + '))';
						END
						ELSE
						BEGIN
							SET @sTemp = CASE
									WHEN @fAddedToPositionString = 0 THEN ''
									ELSE ' OR '
								END + 
								'(' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ' + convert(varchar(200), @dblNumValue) + ') OR ('  + @sOrderItem + ' IS NULL)))';
						END

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ' + convert(varchar(MAX), @dblNumValue) + ')';
					END
				END
			END

			IF @iDataType = 61	/* datetime */
			BEGIN
				SET @sLongCommand = 'SELECT TOP 1 @dtValue = ' + @sOrderItem +
					' FROM ' + @sFromSQL +
					' WHERE ' + @sRealSource + '.id = ' + convert(varchar(100), @iRecordID);
				SET @sSubParamDefinition = N'@dtValue datetime OUTPUT';
				EXEC sp_executesql @sLongCommand,  @sSubParamDefinition, @dtDateValue OUTPUT;

				IF @dtDateValue IS NULL
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = '(NOT ' + @sOrderItem + ' IS NULL)';
							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END

						SET @sPreviousWhere = '(' + @sOrderItem + ' IS NULL)';
					END
					ELSE
					BEGIN
						IF @fDescending = 1 
						BEGIN
							SET @sTemp = CASE
								WHEN @fAddedToPositionString = 0 THEN ''
								ELSE ' OR '
							END + 
							'(' + @sPreviousWhere + ' AND (NOT' + @sOrderItem + ' IS NULL))';

							SET @sPositionCommand = @sPositionCommand + @sTemp;
							SET @fAddedToPositionString = 1;
						END

						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' IS NULL)';
					END
				END
				ELSE
				BEGIN
					IF LEN(@sPreviousWhere) = 0
					BEGIN
						IF @fDescending = 1
							SET @sTemp = '(' + @sOrderItem + ' > ''' + convert(varchar(50), @dtDateValue, 121)  + ''')';
						ELSE
							SET @sTemp = '((' + @sOrderItem + ' < ''' + convert(varchar(50), @dtDateValue, 121)  + ''') OR ('  + @sOrderItem + ' IS NULL))';

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = '(' + @sOrderItem + ' = ''' + convert(varchar(50), @dtDateValue, 121) + ''')';
					END
					ELSE
					BEGIN
						IF @fDescending = 1
						BEGIN
							SET @sTemp = CASE
								WHEN @fAddedToPositionString = 0 THEN ''
								ELSE ' OR '
							END + 
							'(' + @sPreviousWhere + ' AND (' + @sOrderItem + ' > ''' + convert(varchar(50), @dtDateValue, 121) + '''))';
						END
						ELSE
						BEGIN
							SET @sTemp = CASE
								WHEN @fAddedToPositionString = 0 THEN ''
								ELSE ' OR '
							END + 
							'(' + @sPreviousWhere + ' AND ((' + @sOrderItem + ' < ''' + convert(varchar(50), @dtDateValue, 121) + ''') OR ('  + @sOrderItem + ' IS NULL)))';
						END

						SET @sPositionCommand = @sPositionCommand + @sTemp;
						SET @fAddedToPositionString = 1;
						SET @sPreviousWhere = @sPreviousWhere + ' AND (' + @sOrderItem + ' = ''' + convert(varchar(50), @dtDateValue, 121) + ''')';
					END
				END
			END

			SET @iLastCharIndex = @iCharIndex;
			SET @iCharIndex = CHARINDEX(', ', @sOrderSQL, @iLastCharIndex + 1);
			SET @sRemainingSQL = SUBSTRING(@sOrderSQL, @iLastCharIndex + 2, len(@sOrderSQL) - @iLastCharIndex);
		END

		/* Add on the ID condition. */
		IF LEN(@sPreviousWhere) = 0
		BEGIN
			SET @sTemp = '((' + @sRealSource + '.id < ' + convert(varchar(255), @iRecordID)  + ') OR ('  + @sRealSource + '.id IS NULL))';
		END
		ELSE
		BEGIN
			SET @sTemp = CASE
				WHEN @fAddedToPositionString = 0 THEN ''
				ELSE ' OR '
			END + 
			'(' + @sPreviousWhere + ' AND ((' + @sRealSource + '.id < ' + convert(varchar(100), @iRecordID) + ') OR ('  + @sRealSource + '.id IS NULL)))';
		END

		SET @sTemp = @sTemp + ')';
		SET @sPositionCommand = @sPositionCommand + @sTemp;
		SET @fAddedToPositionString = 1;

		EXECUTE sp_executeSQL @sPositionCommand;

		set @sPositionSQL = 'SELECT @recordPosition = result FROM ' + @sTempName;
		SET @sPositionParamDefinition = N'@recordPosition integer OUTPUT';
		EXEC sp_executesql @sPositionSQL, @sPositionParamDefinition, @iRecordPosition OUTPUT;
		EXECUTE [dbo].[sp_ASRDropUniqueObject] @sTempName, 3;
		SET @iRecordPosition = @iRecordPosition + 1;
	END

	/* Set the output parameter values. */
	SET @piRecordID = @iRecordID;
	SET @piRecordPosition = @iRecordPosition;

	/* Return the required record. */
	SET @sExecuteSQL = 'SELECT ' + @sSelectSQL
		+ ' FROM ' + @sFromSQL
		+ ' WHERE ' + @sRealSource + '.id = '
		+ convert(varchar(10), @iRecordID);
	EXEC sp_executesql @sExecuteSQL;

END