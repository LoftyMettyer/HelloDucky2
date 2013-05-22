CREATE PROCEDURE [dbo].[sp_ASRIntCalcDefaults] (
	@piRecordCount			integer OUTPUT,
	@psFromDef 				varchar(MAX),
	@psFilterDef 			varchar(MAX),
	@piTableID				integer,
	@piParentTableID		integer,
	@piParentRecordID		integer,
	@psDefaultCalcColumns	varchar(MAX),
	@psDecimalSeparator		varchar(255),
	@psLocaleDateFormat		varchar(255)
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @iRecordCount 	integer,
		@sCommand			nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@sColumns			varchar(MAX),
		@iID				integer,
		@iDataType			integer,
		@iSize				integer,
		@iDecimals			integer,
		@iDfltExprID		integer,
		@fOneColumnDone		bit,
		@iCount				integer,
		@fOK				bit,
		@iTableID			integer,
		@sCharResult 		varchar(MAX),
		@dblNumericResult 	float,
		@iIntegerResult 	integer,
		@dtDateResult 		datetime,
		@fLogicResult 		bit,
		@sTempTableName		sysname,
		@sTemp 				sysname,
		@iLoop 				integer,
		@sTempDate			varchar(10),
		@iIndex1			integer,
		@iIndex2			integer,
		@iIndex3			integer,
		@iColumnID			integer,
		@iOperatorID		integer,
		@sValue				varchar(MAX),
		@sFilterSQL			nvarchar(MAX),
		@sSubFilterSQL		nvarchar(MAX),
		@sColumnName 		sysname,
		@sFromSQL			nvarchar(MAX),
		@sRealSource		sysname,
		@sRealSourceAlias	varchar(MAX),
		@sTableViewName		sysname,
		@iJoinTableID		integer,
		@sColumnsDone		varchar(MAX);
		
	SET @fOneColumnDone = 0;
	SET @fOK = 1;
	SET @sFilterSQL = '';

	SET @iIndex1 = charindex(char(9), @psFromDef);
	SET @sRealSource = replace(LEFT(@psFromDef, @iIndex1-1), '''', '''''');
	SET @sRealSourceAlias = 'RS';
	SET @sFromSQL = @sRealSource + ' ' + @sRealSourceAlias + ' ';
	SET @psFromDef = SUBSTRING(@psFromDef, @iIndex1+1, LEN(@psFromDef) - @iIndex1);

	WHILE charindex(char(9), @psFromDef) > 0
	BEGIN
		SET @iIndex1 = charindex(char(9), @psFromDef);
		SET @iIndex2 = charindex(char(9), @psFromDef, @iIndex1+1);
				
		SET @sTableViewName = replace(LEFT(@psFromDef, @iIndex1-1), '''', '''''');
		SET @iJoinTableID = convert(integer, SUBSTRING(@psFromDef, @iIndex1+1, @iIndex2-@iIndex1-1));
				
		SET @psFromDef = SUBSTRING(@psFromDef, @iIndex2+1, LEN(@psFromDef) - @iIndex2);

		SET @sFromSQL = @sFromSQL + 
			' LEFT OUTER JOIN ' + @sTableViewName + ' ON ' + @sRealSourceAlias + '.ID_' + convert(varchar(100), @iJoinTableID) + ' = ' + @sTableViewName + '.ID';
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
							
			SET @sColumnName = @sRealSourceAlias + '.' + @sColumnName;

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
	
	/* Get the record count of the current recordset. */
	SET @sCommand = 'SELECT @recordCount = COUNT(' + @sRealSourceAlias + '.id)' +
		' FROM ' + @sFromSQL;

	IF @piParentTableID > 0
	BEGIN
		SET @sCommand = @sCommand +
			' WHERE ' + @sRealSourceAlias + '.id_' + convert(varchar(100), @piParentTableID) + ' = ' + convert(varchar(100), @piParentRecordID);
		
		IF len(@sFilterSQL) > 0
		BEGIN
			SET @sCommand = @sCommand +	' AND ' + @sFilterSQL;
		END
	END
	ELSE
	BEGIN
		IF len(@sFilterSQL) > 0
		BEGIN
			SET @sCommand = @sCommand +	' WHERE ' + @sFilterSQL;
		END
	END

	SET @sParamDefinition = N'@recordCount integer OUTPUT';
	EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordCount OUTPUT;
	SET @piRecordCount = @iRecordCount;

	/* Get the default values for the given columns. */
	SET @sColumns = @psDefaultCalcColumns;
	SET @sColumnsDone = ',';
	WHILE len(@sColumns) > 0
	BEGIN
		IF CHARINDEX(',', @sColumns) > 0
		BEGIN
			SET @iID = convert(integer, left(@sColumns, CHARINDEX(',', @sColumns) - 1));
			SET @sColumns = substring(@sColumns, CHARINDEX(',', @sColumns) + 1, len(@sColumns));
		END
		ELSE
		BEGIN
			SET @iID = convert(integer, @sColumns);
			SET @sColumns = '';
		END

		/* Check the column has not already been done. */
		IF CHARINDEX(',' + convert(varchar(MAX), @iID) + ',', @sColumnsDone) > 0
		BEGIN
			/* Column already been done. */
			SET @iID = 0;
		END
		ELSE
		BEGIN
			/* Column NOT already been done. */
			SET @sColumnsDone = @sColumnsDone + convert(varchar(MAX), @iID) + ',';
		END

		IF @iID > 0 			
		BEGIN
			/* Get the data type and size of the column. */
			SELECT @iDataType = dataType, 
				@iSize = size, 
				@iDecimals = decimals,
				@iDfltExprID = dfltValueExprID
			FROM ASRSysColumns
			WHERE columnID = @iID;

			/* Check the default expression stored procedure exists. */
			SET @sCommand = 'SELECT @count = COUNT(*)' +
				' FROM sysobjects' +
				' WHERE id = object_id(N''sp_ASRDfltExpr_' + convert(varchar(100), @iDfltExprID) + ''')' +
				' AND OBJECTPROPERTY(id, N''IsProcedure'') = 1';
			SET @sParamDefinition = N'@count integer OUTPUT';
			EXEC sp_executesql @sCommand,  @sParamDefinition, @iCount OUTPUT;

			IF @iCount > 0 
			BEGIN
				SET @sCommand = 'exec sp_ASRDfltExpr_' + convert(varchar(100), @iDfltExprID) + ' @result output';	
				SET @fOK = 0;

				IF @iDataType = -7 /* Logic columns. */
				BEGIN
					SET @sParamDefinition = N'@result bit OUTPUT';
					SET @fOK = 1;
				END
          
				IF @iDataType = 2 /* Numeric columns. */
				BEGIN
					SET @sParamDefinition = N'@result float OUTPUT';
					SET @fOK = 1;
				END
          
				IF @iDataType = 4 /* Integer columns. */
				BEGIN
					SET @sParamDefinition = N'@result integer OUTPUT';
					SET @fOK = 1;
				END
          
				IF @iDataType = 11 /* Date columns. */
				BEGIN
					SET @sParamDefinition = N'@result datetime OUTPUT';
					SET @fOK = 1;
				END
          
				IF @iDataType = 12 /* Character columns. */
				BEGIN
					SET @sParamDefinition = N'@result varchar(MAX) OUTPUT';
					SET @fOK = 1;
				END
          
				IF @iDataType = -1 /* Working Pattern columns. */
				BEGIN
					SET @sParamDefinition = N'@result varchar(14) OUTPUT';
					SET @fOK = 1;
				END

				IF @fOK = 1
				BEGIN
 					/* Append the parent table ID parameters. */
					DECLARE parentsCursor CURSOR LOCAL FAST_FORWARD FOR 
						SELECT parentID

						FROM ASRSysRelations
						WHERE childID = @piTableID
						ORDER BY parentID;
					OPEN parentsCursor;
					FETCH NEXT FROM parentsCursor INTO @iTableID;
					WHILE (@@fetch_status = 0)
					BEGIN
						IF @iTableID = @piParentTableID
						BEGIN
							SET @sCommand = @sCommand + ', ' + convert(varchar(100), @piParentRecordID);
						END
						ELSE
						BEGIN
							SET @sCommand = @sCommand + ', 0';
						END

						FETCH NEXT FROM parentsCursor INTO @iTableID;
					END
					CLOSE parentsCursor;
					DEALLOCATE parentsCursor;

					IF @iDataType = -7 /* Logic columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @fLogicResult OUTPUT;
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = '';
							SET @iLoop = 1;
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = 'tmpDefaultValues_' + convert(varchar(100), @iLoop);

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp;

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp;
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1;
								END
							END

							SET @sCommand = 'CREATE TABLE ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + '] bit NULL)';
							EXEC sp_executesql @sCommand;

							SET @sCommand = 'INSERT INTO ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + ']) VALUES (@newValue)';
							SET @sParamDefinition = N'@newValue bit';
							EXEC sp_executesql @sCommand,  @sParamDefinition, @fLogicResult;
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = 'ALTER TABLE ' + @sTempTableName +' ADD [' + convert(varchar(100), @iID) + '] bit NULL';
							EXEC sp_executesql @sCommand;
							SET @sCommand = 'UPDATE ' + @sTempTableName +' SET [' + convert(varchar(100), @iID) + '] = ' + convert(nvarchar(MAX), @fLogicResult);
							EXEC sp_executesql @sCommand;
						END
					END
          
					IF @iDataType = 2 /* Numeric columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @dblNumericResult OUTPUT;
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = '';
							SET @iLoop = 1;
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = 'tmpDefaultValues_' + convert(varchar(100), @iLoop);

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp;

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp;
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1;
								END
							END

							SET @sCommand = 'CREATE TABLE ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + '] float NULL)';
							EXEC sp_executesql @sCommand;

							SET @sCommand = 'INSERT INTO ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + ']) VALUES (@newValue)';
							SET @sParamDefinition = N'@newValue float';
							EXEC sp_executesql @sCommand,  @sParamDefinition, @dblNumericResult;
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = 'ALTER TABLE ' + @sTempTableName +' ADD [' + convert(varchar(100), @iID) + '] float NULL';
							EXEC sp_executesql @sCommand;
							SET @sCommand = 'UPDATE ' + @sTempTableName +' SET [' + convert(varchar(100), @iID) + '] = ' + convert(nvarchar(MAX), @dblNumericResult);
							EXEC sp_executesql @sCommand;
						END
					END
          
					IF @iDataType = 4 /* Integer columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @iIntegerResult OUTPUT;
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = '';
							SET @iLoop = 1;
							WHILE len(@sTempTableName) = 0

							BEGIN
								SET @sTemp = 'tmpDefaultValues_' + convert(varchar(100), @iLoop);

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp;

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp;
								END


								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1;
								END
							END

							SET @sCommand = 'CREATE TABLE ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + '] integer NULL)';
							EXEC sp_executesql @sCommand;

							SET @sCommand = 'INSERT INTO ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + ']) VALUES (@newValue)';
							SET @sParamDefinition = N'@newValue integer';
							EXEC sp_executesql @sCommand,  @sParamDefinition, @iIntegerResult;
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = 'ALTER TABLE ' + @sTempTableName +' ADD [' + convert(varchar(100), @iID) + '] integer NULL';
							EXEC sp_executesql @sCommand;
							SET @sCommand = 'UPDATE ' + @sTempTableName +' SET [' + convert(varchar(100), @iID) + '] = ' + convert(nvarchar(MAX), @iIntegerResult);
							EXEC sp_executesql @sCommand;
						END
					END
          
					IF @iDataType = 11 /* Date columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @dtDateResult OUTPUT;
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = '';
							SET @iLoop = 1;
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = 'tmpDefaultValues_' + convert(varchar(100), @iLoop);

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp;

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp;
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1;
								END
							END

							SET @sCommand = 'CREATE TABLE ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + '] varchar(10) NULL)';
							EXEC sp_executesql @sCommand;

							SET @sCommand = 'INSERT INTO ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + ']) VALUES (@newValue)';
							SET @sParamDefinition = N'@newValue varchar(10)';

							SET @sTempDate = convert(varchar(10), @dtDateResult, 101);
							EXEC sp_executesql @sCommand,  @sParamDefinition, @sTempDate;
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = 'ALTER TABLE ' + @sTempTableName +' ADD [' + convert(varchar(100), @iID) + '] varchar(10) NULL';
							EXEC sp_executesql @sCommand;
							SET @sCommand = 'UPDATE ' + @sTempTableName +' SET [' + convert(varchar(100), @iID) + '] = ''' + convert(nvarchar(MAX), @dtDateResult, 101) + '''';
							EXEC sp_executesql @sCommand;
						END
					END
          	
					IF @iDataType = 12 /* Character columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @sCharResult OUTPUT;
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = '';
							SET @iLoop = 1;
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = 'tmpDefaultValues_' + convert(varchar(100), @iLoop);

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp;

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp;
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1;
								END
							END

							SET @sCommand = 'CREATE TABLE ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + '] varchar(MAX) NULL)';
							EXEC sp_executesql @sCommand;
							SET @sCommand = 'INSERT INTO ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + ']) VALUES (@newValue)';
							SET @sParamDefinition = N'@newValue varchar(MAX)';
							EXEC sp_executesql @sCommand,  @sParamDefinition, @sCharResult;
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = 'ALTER TABLE ' + @sTempTableName +' ADD [' + convert(varchar(100), @iID) + '] varchar(MAX) NULL';
							EXEC sp_executesql @sCommand;
							SET @sCommand = 'UPDATE ' + @sTempTableName +' SET [' + convert(varchar(100), @iID) + '] = ''' + REPLACE(convert(nvarchar(MAX), @sCharResult), '''', '''''') + '''';
							EXEC sp_executesql @sCommand;
						END
					END
          
					IF @iDataType = -1 /* Working Pattern columns. */
					BEGIN
						EXEC sp_executesql @sCommand,  @sParamDefinition, @sCharResult OUTPUT;
						IF @fOneColumnDone = 0
						BEGIN
							/* Create the temp table to hold the default values. */
							SET @sTempTableName = '';
							SET @iLoop = 1;
							WHILE len(@sTempTableName) = 0
							BEGIN
								SET @sTemp = 'tmpDefaultValues_' + convert(varchar(100), @iLoop);

								SELECT @icount = COUNT(*)
								FROM sysobjects
								WHERE name = @sTemp;

								IF @iCount = 0
								BEGIN
									SET @sTempTableName = @sTemp;
								END
								ELSE
								BEGIN
									SET @iLoop = @iLoop + 1;
								END
							END

							SET @sCommand = 'CREATE TABLE ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + '] varchar(MAX) NULL)';
							EXEC sp_executesql @sCommand;

							SET @sCommand = 'INSERT INTO ' + @sTempTableName +' ([' + convert(varchar(100), @iID) + ']) VALUES (@newValue)';
							SET @sParamDefinition = N'@newValue varchar(MAX)';
							EXEC sp_executesql @sCommand,  @sParamDefinition, @sCharResult;
						END
						ELSE
						BEGIN
							/* Alter the temp table. */
							SET @sCommand = 'ALTER TABLE ' + @sTempTableName +' ADD [' + convert(varchar(100), @iID) + '] varchar(MAX) NULL';
							EXEC sp_executesql @sCommand;
							SET @sCommand = 'UPDATE ' + @sTempTableName +' SET [' + convert(varchar(100), @iID) + '] = ''' + REPLACE(convert(nvarchar(MAX), @sCharResult), '''', '''''') + '''';
							EXEC sp_executesql @sCommand;
						END
					END

					SET @fOneColumnDone = 1;
				END
			END
		END
	END

	IF @fOneColumnDone > 0
	BEGIN
		SET @sCommand = 'SELECT * FROM ' + @sTempTableName;
		EXEC sp_executesql @sCommand;

		SET @sCommand = 'DROP TABLE ' + @sTempTableName;
		EXEC sp_executesql @sCommand;
	END
END