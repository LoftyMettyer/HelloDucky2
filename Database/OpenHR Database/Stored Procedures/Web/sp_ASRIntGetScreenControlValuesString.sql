CREATE PROCEDURE [dbo].[sp_ASRIntGetScreenControlValuesString] (
	@plngScreenID 	int)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the column control values in the given screen. */
	DECLARE @lngColumnID	integer,
		@sValue				varchar(MAX),
		@lngLastColumnID	integer,
		@sDefinition		varchar(MAX);

	DECLARE @valuesInfo TABLE (valueDefinition varchar(MAX));

	SET @lngLastColumnID = 0;
	SET @sDefinition = '';

	DECLARE values_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysColumnControlValues.columnID, ASRSysColumnControlValues.value
		FROM ASRSysColumnControlValues
		WHERE ASRSysColumnControlValues.columnID IN (
			SELECT ASRSysControls.columnID
			FROM ASRSysControls
			WHERE ASRSysControls.screenID = @plngScreenID)
		ORDER BY ASRSysColumnControlValues.columnID, ASRSysColumnControlValues.sequence;
				
	OPEN values_cursor;
	FETCH NEXT FROM values_cursor INTO @lngColumnID, @sValue;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @lngLastColumnID <> @lngColumnID
		BEGIN
			IF @lngLastColumnID <> 0
			BEGIN
				INSERT INTO @valuesInfo (valueDefinition) VALUES(@sDefinition);
			END

			SET @sDefinition = convert(varchar(MAX), @lngColumnID) + char(9) + case when @sValue IS null then '' else @sValue end;
		END
		ELSE
		BEGIN
			SET @sDefinition = @sDefinition + char(9) + case when @sValue IS null then '' else @sValue end;
		END

		SET @lngLastColumnID = @lngColumnID;
		FETCH NEXT FROM values_cursor INTO @lngColumnID, @sValue;
	END
	
	CLOSE values_cursor;
	DEALLOCATE values_cursor;

	/* Do the last row. */
	IF @lngLastColumnID <> 0
	BEGIN
		INSERT INTO @valuesInfo (valueDefinition) VALUES(@sDefinition);
	END

	SELECT * FROM @valuesInfo;
END