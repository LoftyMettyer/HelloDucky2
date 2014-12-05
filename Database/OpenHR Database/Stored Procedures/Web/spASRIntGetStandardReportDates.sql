CREATE PROCEDURE [dbo].[spASRIntGetStandardReportDates] (
	@piReportType 		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the prompted values for the given utililty. */
	DECLARE	@sComponents			varchar(MAX),
			@vDateID				varchar(100),
			@iStartDateID			integer,
			@iEndDateID				integer,
			@iStartDateComponentID	integer,
			@iEndDateComponentID	integer;

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
		fieldColumnID		integer,
		StartEndType		varchar(5)
	)

	-- Absence Breakdown	
	IF @piReportType = 15
	BEGIN
		EXEC dbo.spASRIntGetSetting 'AbsenceBreakdown', 'Start Date', 0, 0, @vDateID OUTPUT;
		SET @iStartDateID = convert(integer,@vDateID);

		EXEC dbo.spASRIntGetSetting 'AbsenceBreakdown', 'End Date', 0, 0, @vDateID OUTPUT;
		SET @iEndDateID = convert(integer,@vDateID);
	END

	-- Bradford Factor
	IF @piReportType = 16
	BEGIN
		EXEC dbo.spASRIntGetSetting 'BradfordFactor', 'Start Date', 0, 0, @vDateID OUTPUT;
		SET @iStartDateID = convert(integer,@vDateID);

		EXEC dbo.spASRIntGetSetting 'BradfordFactor', 'End Date', 0, 0, @vDateID OUTPUT;
		SET @iEndDateID = convert(integer,@vDateID);
	END

	/* Start Date prompted value. */		
	IF @iStartDateID > 0
	BEGIN
		EXEC sp_ASRIntGetFilterPromptedValues @iStartDateID, @sComponents OUTPUT
		SET @iStartDateComponentID = @sComponents

		INSERT INTO @promptedValues
			(componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID,StartEndType)
		(SELECT componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID,'start'
			FROM ASRSysExprComponents
			WHERE componentID = @iStartDateComponentID)
	END

	/* End Date prompted value. */
	IF @iEndDateID > 0
	BEGIN
		EXEC sp_ASRIntGetFilterPromptedValues @iEndDateID, @sComponents OUTPUT
		SET @iEndDateComponentID = @sComponents

		INSERT INTO @promptedValues
			(componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID,StartEndType)
		(SELECT componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID,'end'
			FROM ASRSysExprComponents
			WHERE componentID = @iEndDateComponentID)
	END


	SELECT DISTINCT * 
	FROM @promptedValues
	ORDER BY startEndType DESC;
	
END