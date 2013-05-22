CREATE PROCEDURE [dbo].[sp_ASRIntGetFilterPromptedValuesRecordset] (
	@piFilterID 		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the prompted values for the given filter. */
	DECLARE	@iComponentID		integer,
			@sComponents		varchar(MAX),
			@sAllComponents		varchar(MAX),
			@iIndex				integer;
	
	SET @sAllComponents = '';

	/* Create a temp table to hold the propmted value details. */
	DECLARE @promptedValues TABLE(
		componentID			integer,
		promptDescription	varchar(255),
		valueType			integer,
		promptMask			varchar(255),
		promptSize			integer,
		promptDecimals		integer,
		valueCharacter		varchar(255),
		valueNumeric		float,
		valueLogic			bit,
		valueDate			datetime,
		promptDateType		integer, 
		fieldColumnID		integer);
	
	EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @piFilterID, @sComponents OUTPUT;

	IF LEN(@sComponents) > 0
	BEGIN
		SET @sAllComponents = @sAllComponents + 
			CASE
				WHEN LEN(@sAllComponents) > 0 THEN ','
				ELSE ''
			END + 
			@sComponents;
	END


	/* We now have a string of all of the prompted value components that are used in the filter. */
	WHILE LEN(@sAllComponents) > 0 
	BEGIN
		/* Get the individual component IDs from the comma-delimited string. */
		SET @iIndex = CHARINDEX(',', @sAllComponents);

		IF @iIndex > 0 
		BEGIN
			SET @iComponentID = convert(integer, SUBSTRING(@sAllComponents, 1, @iIndex - 1));
			SET @sAllComponents = SUBSTRING(@sAllComponents, @iIndex + 1, LEN(@sAllComponents) - @iIndex);
		END
		ELSE
		BEGIN
			/* No comma, must be dealing with the last component in the list. */
			SET @iComponentID = convert(integer, @sAllComponents);
			SET @sAllComponents = '';
		END

		/* Get the parameters of the prompted values. */
		INSERT INTO @promptedValues
			(componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID)
		(SELECT componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID
			FROM ASRSysExprComponents
			WHERE componentID = @iComponentID);
	END

	SELECT DISTINCT * 
		FROM @promptedValues
		ORDER BY promptDescription;
END