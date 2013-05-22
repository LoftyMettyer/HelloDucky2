CREATE PROCEDURE [dbo].[spASRIntUpdateRecord]
(
	@piResult		integer	OUTPUT,	/* Output variable to hold the result. */
	@psUpdateDef	varchar(MAX),	/* Update definition to update the record. */
	@piTableID		integer,		/* TableID being updated. */
	@psRealSource	sysname,		/* RealSource being updated. */
	@piID			integer,		/* ID the record being updated. */
	@piTimestamp	integer			/* Original timestamp of the record being updated. */
)
AS
BEGIN
	/* Return 0 if the record was OK to update. */
	/* Return 1 if the record has been amended AND is still in the given table/view. */
	/* Return 2 if the record has been amended AND is no longer in the given table/view. */
	/* Return 3 if the record has been deleted from the table. */
	SET NOCOUNT ON;

	DECLARE
		@iCurrentTimestamp	integer,
		@sSQL				nvarchar(MAX),
		@iCount				integer,
		@sUpdateString		nvarchar(MAX),
		@sTempString		varchar(MAX),
		@iCounter			integer,
		@iIndex1			integer,
		@iIndex2			integer,
		@sColumnID			varchar(255),
		@sValue				varchar(MAX),
		@iDataType			integer,
		@sColumnName		varchar(255),
		@sMask				varchar(MAX),
		@iOLEType			integer;

	-- Clean the input string parameters.
	IF len(@psRealsource) > 0 SET @psRealsource = replace(@psRealsource, '''', '''''');

	SET @piResult = 0;
	SET @sUpdateString = 'UPDATE ' + convert(varchar(255), @psRealSource) + ' SET ';
	SET @iCounter = 0;

	-- Get status of amended record
	EXEC dbo.sp_ASRRecordAmended @piResult OUTPUT,
	    @piTableID,
		@psRealSource,
		@piID,
		@piTimestamp;

	IF @piResult = 0
	BEGIN
		WHILE charindex(CHAR(9), @psUpdateDef) > 0
		BEGIN
			SET @iIndex1 = charindex(CHAR(9), @psUpdateDef);
			SET @iIndex2 = charindex(CHAR(9), @psUpdateDef, @iIndex1+1);

			SET @sColumnID = replace(LEFT(@psUpdateDef, @iIndex1-1), '''', '''''');
			SET @sValue = replace(SUBSTRING(@psUpdateDef, @iIndex1+1, @iIndex2-@iIndex1-1), '''', '''''');

			IF LEFT(@sColumnID, 3) = 'ID_'
			BEGIN
				SET @sColumnName = @sColumnID;
			END
			ELSE
			BEGIN
				SELECT @sColumnName = ASRSysColumns.columnName,
					@iDataType = ASRSysColumns.dataType,
					@sMask = ASRSysColumns.mask
				FROM ASRSysColumns
				WHERE ASRSysColumns.columnID = convert(integer, @sColumnID);

				-- Date
				IF (@iDataType = 11 AND @sValue <> 'null') SET @sValue = '''' + @sValue + '''';

				-- Character
				IF (@iDataType = 12 AND (LEN(@sMask) = 0 OR @sValue <> 'null')) SET @sValue = '''' + @sValue + '''';

				-- WorkingPattern
				IF (@iDataType = -1) SET @sValue = '''' + @sValue + '''';

				-- Photo / OLE
				IF (@iDataType = -3 OR @iDataType = -4)
				BEGIN
					SET @iOLEType = convert(integer, LEFT(@sValue, 1));
					SET @sValue = SUBSTRING(@sValue, 2, LEN(@sValue) - 1);
					IF (@iOLEType < 2) SET @sValue = '''' + @sValue + '''';
				END
			END

			SET @sTempString =
				CASE
					WHEN @iCounter > 0 THEN ','
					ELSE ''
				END
				+ convert(varchar(255), @sColumnName) + ' = ' + convert(varchar(MAX), @sValue);

			SET @sUpdateString = @sUpdateString + @sTempString;
			SET @iCounter = @iCounter + 1;
			SET @psUpdateDef = SUBSTRING(@psUpdateDef, @iIndex2+1, LEN(@psUpdateDef) - @iIndex2);
		END

		SET @sUpdateString = @sUpdateString + ' WHERE id = ' + convert(varchar(255), @piID);

		-- Run the constructed SQL UPDATE string.
		EXEC sp_executeSQL @sUpdateString;
	END
END