CREATE PROCEDURE [dbo].[sp_ASRFn_GetFieldFromDatabase] (
	@psCharResult		varchar(255) OUTPUT,
	@pfBitResult		bit	OUTPUT,
	@pfltNumResult		float OUTPUT,
	@pdtDateResult		datetime OUTPUT,
	@piSearchColumnID	int,
	@psCharSearchValue	varchar(255),
	@pfBitSearchValue	bit,
	@pfltNumSearchValue	float,
	@pdtDateSearchValue	datetime,
	@piReturnColumnID	int)
AS
BEGIN
	DECLARE @sSearchColumnName	sysname,
		@sSearchTableName		sysname,
		@sReturnColumnName		sysname,
		@sReturnTableName		sysname,
		@iSearchColumnType		int,
		@iReturnColumnType		int,
		@sCommandString			nvarchar(MAX),
		@sReturnString			nvarchar(MAX),
		@sParamDefinition		nvarchar(500),
		@sNewCharSearchValue	varchar(MAX),
		@iCharacterIndex		int,
		@iStringLength			int,
		@sCurrentChar 			varchar(1)

	/* Replace any single quote characters in the character search string 
	with two single quote characters so that the SQL Select string which is 
	constructed below is still valid for execution. */
	SET @sNewCharSearchValue = '';
	SET @iCharacterIndex = 0;
	SET @iStringLength = LEN(@psCharSearchValue);

	WHILE @iCharacterIndex < @iStringLength
	BEGIN
		SET @iCharacterIndex = @iCharacterIndex + 1;
		SET @sCurrentChar = SUBSTRING(@psCharSearchValue, @iCharacterIndex, 1);
		SET @sNewCharSearchValue = @sNewCharSearchValue + @sCurrentChar;
	
		IF @sCurrentChar = ''''
		BEGIN
			SET @sNewCharSearchValue = @sNewCharSearchValue + @sCurrentChar;
		END
	END

	SET @psCharSearchValue = @sNewCharSearchValue;

	/* Get the name of the search column. */
	SELECT @sSearchColumnName = ASRSysColumns.columnName, 
		@sSearchTableName = ASRSysTables.tableName, 
		@iSearchColumnType = ASRSysColumns.dataType
	FROM ASRSysColumns
	JOIN ASRSysTables 
		ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysColumns.columnID = @piSearchColumnID;

	/* Get the name of the return column. */
	SELECT @sReturnColumnName = ASRSysColumns.columnName, 
		@sReturnTableName = ASRSysTables.tableName, 
		@iReturnColumnType = ASRSysColumns.dataType
	FROM ASRSysColumns
	JOIN ASRSysTables 
		ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysColumns.columnID = @piReturnColumnID;

	IF (NOT @sSearchColumnName IS NULL) 
		AND (NOT @sSearchTableName IS NULL) 
		AND(NOT @sReturnColumnName IS NULL) 
		AND (NOT @sReturnTableName IS NULL)
		AND ((@iSearchColumnType = 12) OR (@iSearchColumnType = -7) OR (@iSearchColumnType = 4) OR (@iSearchColumnType = 2) OR (@iSearchColumnType = 11)) 
		AND ((@iReturnColumnType = 12) OR (@iReturnColumnType = -7) OR (@iReturnColumnType = 4) OR (@iReturnColumnType = 2) OR (@iReturnColumnType = 11)) 
		AND (@sSearchTableName = @sReturnTableName)
	BEGIN
		IF @iReturnColumnType = 12 
		BEGIN
			SET @sReturnString = '@charResult';
			SET @sParamDefinition = N'@charResult varchar(255) OUTPUT';
		END

		IF @iReturnColumnType = -7 
		BEGIN
			SET @sReturnString = '@bitResult';
			SET @sParamDefinition = N'@bitResult bit OUTPUT';
		END

		IF (@iReturnColumnType = 4) OR (@iReturnColumnType = 2) 
		BEGIN
			SET @sReturnString = '@numResult';
			SET @sParamDefinition = N'@numResult float OUTPUT';
		END

		IF @iReturnColumnType = 11 
		BEGIN
			SET @sReturnString = '@datetimeResult';
			SET @sParamDefinition = N'@dateResult datetime OUTPUT';
		END

		IF @iSearchColumnType = 12 
		BEGIN
			SET @sCommandString = 'SELECT ' + @sReturnString + ' = ' + @sReturnColumnName + ' FROM ' + @sReturnTableName + ' WHERE ' + @sSearchColumnName + ' = ''' + @psCharSearchValue + '''';
		END

		IF @iSearchColumnType = -7 
		BEGIN
			SET @sCommandString = 'SELECT  ' + @sReturnString + ' = ' + @sReturnColumnName + ' FROM ' + @sReturnTableName + ' WHERE ' + @sSearchColumnName + ' = ' + convert(varchar(MAX), @pfBitSearchValue);
		END

		IF (@iSearchColumnType = 4) OR (@iSearchColumnType = 2) 
		BEGIN
			SET @sCommandString = 'SELECT  ' + @sReturnString + ' = ' + @sReturnColumnName + ' FROM ' + @sReturnTableName + ' WHERE ' + @sSearchColumnName + ' = ' + convert(varchar(MAX), @pfltNumSearchValue)
		END
	
		IF @iSearchColumnType = 11 
		BEGIN
			SET @sCommandString = 'SELECT  ' + @sReturnString + ' = ' + @sReturnColumnName + ' FROM ' + @sReturnTableName + ' WHERE ' + @sSearchColumnName + ' = ''' + convert(varchar(MAX), @pdtDateSearchValue, 101) + ''''
		END
		IF @iReturnColumnType = 12 EXECUTE sp_executesql @sCommandString, @sParamDefinition, @psCharResult OUTPUT;
		IF @iReturnColumnType = -7 EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pfBitResult OUTPUT;
		IF (@iReturnColumnType = 4) OR (@iReturnColumnType = 2) EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pfltNumResult OUTPUT;
		IF @iReturnColumnType = 11 EXECUTE sp_executesql @sCommandString, @sParamDefinition, @pdtDateResult OUTPUT;
	END

	/* Return the result. */
	IF @iReturnColumnType = 12 SELECT @psCharResult AS result;
	IF @iReturnColumnType = -7 SELECT @pfBitResult AS result;
	IF ((@iReturnColumnType = 4) OR (@iReturnColumnType = 2)) SELECT @pfltNumResult AS result;
	IF @iReturnColumnType = 11 SELECT @pdtDateResult AS result;
END
