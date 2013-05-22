CREATE PROCEDURE [dbo].[spASRIntInsertNewRecord]
(
	@piNewRecordID	integer	OUTPUT,	/* Output variable to hold the new record ID. */
	@psInsertDef	nvarchar(MAX)	/* SQL Insert string to insert the new record. */
)
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE
		@sTempString	nvarchar(MAX),
		@sInsertString	nvarchar(MAX),
		@iTemp			integer,
		@iCounter		integer,
		@iIndex1		integer,
		@iIndex2		integer,
		@iIndex3		integer,
		@sColumnID		varchar(255),
		@sValue			varchar(MAX),
		@sColumnList	varchar(MAX),
		@sValueList		varchar(MAX),
		@iCopiedRecordID	integer,
		@iDataType		integer,
		@sColumnName	varchar(255),
		@sRealSource	sysname,
		@sMask			varchar(255),
		@iOLEType		integer,
		@fCopyImageData	bit,
		@tablename		varchar(255);

	SET @sColumnList = '';
	SET @sValueList = '';

	SET @iIndex1 = charindex(CHAR(9), @psInsertDef);
	SET @iIndex2 = charindex(CHAR(9), @psInsertDef, @iIndex1+1);
	SET @iIndex3 = charindex(CHAR(9), @psInsertDef, @iIndex2+1);

	SET @sRealSource = replace(LEFT(@psInsertDef, @iIndex1-1), '''', '''''');
	SET @sValue = replace(SUBSTRING(@psInsertDef, @iIndex1+1, @iIndex2-@iIndex1-1), '''', '''''');
	SET @fCopyImageData = convert(bit, @sValue);
	SET @sValue = replace(SUBSTRING(@psInsertDef, @iIndex2+1, @iIndex3-@iIndex2-1), '''', '''''');
	SET @iCopiedRecordID = convert(integer, @sValue);

	SET @psInsertDef = SUBSTRING(@psInsertDef, @iIndex3+1, LEN(@psInsertDef) - @iIndex3);

	SET @sColumnList = 'INSERT ' + convert(varchar(255), @sRealSource) + ' (';
	SET @sValueList = '';
	SET @iCounter = 0;

	WHILE charindex(CHAR(9), @psInsertDef) > 0
	BEGIN
		SET @iIndex1 = charindex(CHAR(9), @psInsertDef);
		SET @iIndex2 = charindex(CHAR(9), @psInsertDef, @iIndex1+1);

		SET @sColumnID = replace(LEFT(@psInsertDef, @iIndex1-1), '''', '''''');
		SET @sValue = replace(SUBSTRING(@psInsertDef, @iIndex1+1, @iIndex2-@iIndex1-1), '''', '''''');

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
			+ convert(varchar(255), @sColumnName);

		SET @sColumnList = @sColumnList + @sTempString;
		SET @sTempString =
			CASE
				WHEN @iCounter > 0 THEN ','
				ELSE ''
			END
			+ CASE
				WHEN @fCopyImageData = 1 THEN REPLACE(convert(varchar(MAX), @sValue), '''', '''''')
				ELSE convert(varchar(MAX), @sValue)
			END;

		SET @sValueList = @sValueList + @sTempString;
		SET @iCounter = @iCounter + 1;
		SET @psInsertDef = SUBSTRING(@psInsertDef, @iIndex2+1, LEN(@psInsertDef) - @iIndex2);
	END

	IF @fCopyImageData = 1
	BEGIN
		SET @sInsertString = @sColumnList + ')'
			+ ' EXECUTE(''SELECT ' + @sValueList
			+ ' FROM ' + convert(varchar(255), @sRealSource)
			+ ' WHERE id = ' + convert(varchar(255), @iCopiedRecordID) + ''')';
	END
	ELSE
	BEGIN
		SET @sInsertString = @sColumnList + ')' + ' VALUES(' + @sValueList + ')';
	END

	-- Run the constructed SQL INSERT string
	EXECUTE sp_executesql @sInsertString;

	-- Calculate the ID
	SET  @sInsertString = REPLACE(' ' + @sInsertString,' INSERT INTO ','')
	SET  @sInsertString = REPLACE(' ' + @sInsertString,' INSERT ','')
	SET @tablename = SUBSTRING(@sInsertString,0, CHARINDEX('(', @sInsertString));

	SET @sTempString = 'SELECT @ID = MAX(ID) FROM ' + @tablename;
	EXECUTE sp_executesql @sTempString, N'@ID int OUTPUT', @ID = @piNewRecordID OUTPUT;
						  
END