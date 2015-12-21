CREATE PROCEDURE [dbo].[spASRIntGetUtilityPromptedValues] (
	@piUtilType 	integer,
	@piUtilID 		integer,
	@piRecordID 	integer,
	@piMultipleRecords integer = 0
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of the prompted values for the given utililty. */
	DECLARE	@iBaseFilter		integer,
			@iBase2Filter		integer,
			@iParent1Filter		integer,
			@iParent2Filter		integer,
			@iChildFilter		integer,
			@iEventFilter		integer,
			@iStartDateCalc		integer,
			@iEndDateCalc		integer,
			@iDescCalc			integer,
			@iLoop				integer,
			@iFilterID			integer,
			@iCalcID			integer,
			@iComponentID		integer,
			@sComponents		varchar(MAX),
			@sAllComponents		varchar(MAX),
			@iIndex				integer;
	
	SET @sAllComponents = '';

	IF @piRecordID IS null SET @piRecordID = 0;

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


	IF @piUtilType = 1 OR @piUtilType = 35
	BEGIN
		/* Cross Tabs. */
		SELECT @iFilterID = filterid
		FROM ASRSysCrossTab
		WHERE CrossTabID = @piUtilID

		IF (NOT @iFilterID IS NULL) AND (@iFilterID > 0)
		BEGIN
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iFilterID, @sAllComponents OUTPUT
		END		
	END

	IF @piUtilType = 11
	BEGIN
		IF (NOT @piUtilID IS NULL) AND (@piUtilID > 0)
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @piUtilID, @sAllComponents OUTPUT;
	END

	IF @piUtilType = 3
	BEGIN

		SELECT @iBaseFilter = filterID
			FROM [dbo].ASRSysDataTransferName
			WHERE DataTransferID = @piUtilID;

		IF (NOT @iBaseFilter IS NULL) AND (@iBaseFilter > 0)
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iBaseFilter, @sAllComponents OUTPUT;

	END


	IF @piUtilType = 14
	BEGIN

		SELECT @iBaseFilter = Table1filter, @iBase2Filter = Table2Filter
			FROM [dbo].ASRSysMatchReportName
			WHERE MatchReportID = @piUtilID;

		IF (NOT @iBaseFilter IS NULL) AND (@iBaseFilter > 0)
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iBaseFilter, @sAllComponents OUTPUT;

		IF (NOT @iBase2Filter IS NULL) AND (@iBase2Filter > 0)
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iBase2Filter, @sComponents OUTPUT;

		IF LEN(@sComponents) > 0
		BEGIN
			SET @sAllComponents = @sAllComponents + 
				CASE
					WHEN LEN(@sAllComponents) > 0 THEN ','
					ELSE ''
				END + 
				@sComponents
		END

	END

	IF @piUtilType = 38
	BEGIN

		SELECT @iBaseFilter = BaseFilterID, @iBase2Filter = MatchFilterID
			FROM [dbo].ASRSysTalentReports
			WHERE ID = @piUtilID;

		IF (NOT @iBaseFilter IS NULL) AND (@iBaseFilter > 0)
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iBaseFilter, @sAllComponents OUTPUT;

		IF (NOT @iBase2Filter IS NULL) AND (@iBase2Filter > 0)
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iBase2Filter, @sComponents OUTPUT;

		IF LEN(@sComponents) > 0
		BEGIN
			SET @sAllComponents = @sAllComponents + 
				CASE
					WHEN LEN(@sAllComponents) > 0 THEN ','
					ELSE ''
				END + 
				@sComponents
		END

	END


	IF @piUtilType = 15 OR @piUtilType = 16
	BEGIN
		/* Standard report (Absence Calendar or Bradford Factor) */
		IF (NOT @piUtilID IS NULL) AND (@piUtilID > 0)
		BEGIN
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @piUtilID, @sAllComponents OUTPUT
		END		
	END



	IF @piUtilType = 2 OR @piUtilType = 9
	BEGIN

		IF @piUtilType = 2
		BEGIN
			/* Custom Reports. */

			/* Get the IDs of the filters used in the report. */
			SELECT @iBaseFilter = filter, 
				@iParent1Filter = parent1Filter, 
				@iParent2Filter = parent2Filter /*, 
				@iChildFilter = childFilter*/
			FROM [dbo].[ASRSysCustomReportsName]
			WHERE ID = @piUtilID

			IF (@piRecordID <> 0) OR (@piMultipleRecords <> 0)
			BEGIN
				SET @iBaseFilter = 0
			END
			
			/* Get the prompted values used in the Base and Parent table filters. */
			SET @iLoop = 0
			WHILE @iLoop < 3
			BEGIN
				IF @iLoop = 0 SET @iFilterID = @iBaseFilter
				IF @iLoop = 1 SET @iFilterID = @iParent1Filter
				IF @iLoop = 2 SET @iFilterID = @iParent2Filter
				--IF @iLoop = 3 SET @iFilterID = @iChildFilter

				IF (NOT @iFilterID IS NULL) AND (@iFilterID > 0)
				BEGIN
					EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iFilterID, @sComponents OUTPUT

					IF LEN(@sComponents) > 0
					BEGIN
						SET @sAllComponents = @sAllComponents + 
							CASE
								WHEN LEN(@sAllComponents) > 0 THEN ','
								ELSE ''
							END + 
							@sComponents
					END
				END

				SET @iLoop = @iLoop + 1
			END		

			/* Get the promted values used in the Child table filters. */
			DECLARE childs_cursor CURSOR LOCAL FAST_FORWARD FOR
				
			SELECT childFilter
			FROM [dbo].[ASRSysCustomReportsChildDetails]
			WHERE CustomReportID = @piUtilID

			OPEN childs_cursor
			FETCH NEXT FROM childs_cursor INTO @iChildFilter
			WHILE (@@fetch_status = 0)
			BEGIN
				EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iChildFilter, @sComponents OUTPUT
		
				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
				FETCH NEXT FROM childs_cursor INTO @iChildFilter
			END
			CLOSE childs_cursor
			DEALLOCATE childs_cursor


			/* Get the prompted values used in the runtime calcs in the report. */
			DECLARE calcs_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT colExprID 
				FROM [dbo].[ASRSysCustomReportsDetails]
				WHERE customReportID = @piUtilID
					AND type = 'E'

		END

		IF @piUtilType = 9
		BEGIN
			/* Mail Merge. */

			/* Get the IDs of the filters used in the report. */
			SELECT @iFilterID = filterID
			FROM [dbo].[ASRSysMailMergeName]
			WHERE MailMergeID = @piUtilID

			IF (@piRecordID <> 0) OR (@piMultipleRecords <> 0)
			BEGIN
				SET @iFilterID = 0
			END

			IF (NOT @iFilterID IS NULL) AND (@iFilterID > 0)
			BEGIN
				EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iFilterID, @sAllComponents OUTPUT
			END		

			/* Get the prompted values used in the runtime calcs in the report. */
			DECLARE calcs_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT ColumnID 
				FROM [dbo].[ASRSysMailMergeColumns]
				WHERE MailMergeID = @piUtilID
					AND type = 'E'
		END


		OPEN calcs_cursor
		FETCH NEXT FROM calcs_cursor INTO @iCalcID
		WHILE (@@fetch_status = 0)
		BEGIN
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iCalcID, @sComponents OUTPUT
		
			IF LEN(@sComponents) > 0
			BEGIN
				SET @sAllComponents = @sAllComponents + 
					CASE
						WHEN LEN(@sAllComponents) > 0 THEN ','
						ELSE ''
					END + 
					@sComponents
			END
			FETCH NEXT FROM calcs_cursor INTO @iCalcID
		END
		CLOSE calcs_cursor
		DEALLOCATE calcs_cursor
	END

	IF @piUtilType = 17
		BEGIN
			/* Calendar Reports. */

			/* Get the IDs of the filters used in the report. */
			SELECT @iBaseFilter = filter, 
						 @iStartDateCalc = StartDateExpr, 
			 			 @iEndDateCalc = EndDateExpr,
						 @iDescCalc = DescriptionExpr
			FROM ASRSysCalendarReports
			WHERE ID = @piUtilID
				
			IF (@piRecordID = 0) OR (@piMultipleRecords = 0)
			BEGIN
				/* Get the prompted values used in the Base table filter. */
				SET @iFilterID = @iBaseFilter

				IF (NOT @iFilterID IS NULL) AND (@iFilterID > 0)
				BEGIN
					EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iFilterID, @sComponents OUTPUT

					IF LEN(@sComponents) > 0
					BEGIN
						SET @sAllComponents = @sAllComponents + 
							CASE
								WHEN LEN(@sAllComponents) > 0 THEN ','
								ELSE ''
							END + 
							@sComponents
					END
				END
			END
			
			/* Get the prompted values used in the Report Start Date Calculation. */
			SET @iCalcID = @iStartDateCalc

			IF (NOT @iCalcID IS NULL) AND (@iCalcID > 0)
			BEGIN
				EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iCalcID, @sComponents OUTPUT

				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
			END		

			/* Get the prompted values used in the Report End Date Calculation. */
			SET @iCalcID = @iEndDateCalc

			IF (NOT @iCalcID IS NULL) AND (@iCalcID > 0)
			BEGIN
				EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iCalcID, @sComponents OUTPUT

				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
			END		

			/* Get the prompted values used in the Report Description Calculation. */
			SET @iCalcID = @iDescCalc

			IF (NOT @iCalcID IS NULL) AND (@iCalcID > 0)
			BEGIN
				EXEC sp_ASRIntGetFilterPromptedValues @iCalcID, @sComponents OUTPUT

				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
			END		

			/* Get the promted values used in the Event table filters. */
			DECLARE events_cursor CURSOR LOCAL FAST_FORWARD FOR
				
			SELECT FilterID
			FROM ASRSysCalendarReportEvents
			WHERE CalendarReportID = @piUtilID

			OPEN events_cursor
			FETCH NEXT FROM events_cursor INTO @iEventFilter
			WHILE (@@fetch_status = 0)
			BEGIN
				EXEC sp_ASRIntGetFilterPromptedValues @iEventFilter, @sComponents OUTPUT
		
				IF LEN(@sComponents) > 0
				BEGIN
					SET @sAllComponents = @sAllComponents + 
						CASE
							WHEN LEN(@sAllComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents
				END
				FETCH NEXT FROM events_cursor INTO @iEventFilter
			END
			CLOSE events_cursor
			DEALLOCATE events_cursor
			
	END
		
		
	/* We now have a string of all of the prompted value components that are used in the filters and calculations. */
	WHILE LEN(@sAllComponents) > 0 
	BEGIN
		/* Get the individual component IDs from the comma-delimited string. */
		SET @iIndex = CHARINDEX(',', @sAllComponents)

		IF @iIndex > 0 
		BEGIN
			SET @iComponentID = convert(integer, SUBSTRING(@sAllComponents, 1, @iIndex - 1))
			SET @sAllComponents = SUBSTRING(@sAllComponents, @iIndex + 1, LEN(@sAllComponents) - @iIndex)
		END
		ELSE
		BEGIN
			/* No comma, must be dealing with the last component in the list. */
			SET @iComponentID = convert(integer, @sAllComponents)
			SET @sAllComponents = ''
		END

		/* Get the parameters of the prompted values. */
		INSERT INTO @promptedValues
			(componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID)
		(SELECT componentID, promptDescription, valueType, promptMask, promptSize, promptDecimals, valueCharacter, valueNumeric, valueLogic, valueDate, promptDateType, fieldColumnID
			FROM ASRSysExprComponents
			WHERE componentID = @iComponentID)
	END

	SELECT DISTINCT * 
	FROM @promptedValues
END