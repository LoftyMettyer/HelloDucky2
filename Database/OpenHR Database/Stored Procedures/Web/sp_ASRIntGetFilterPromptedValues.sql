CREATE PROCEDURE [dbo].[sp_ASRIntGetFilterPromptedValues] (
	@piFilterID 		integer,
	@psComponents		varchar(MAX)	OUTPUT
)
AS
BEGIN
	/* Return a list of the prompted values in the given filter (and sub-filters). */
	DECLARE	@iComponentID	integer, 
			@iType			integer,
			@sComponents	varchar(MAX),
			@iExprID		integer,
			@iFieldFilterID	integer;

	SET @psComponents = '';

	/* Get the prompted value components, and also the subexpressions (sub-filters and function parameters). */
	DECLARE components_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT componentID, 
			type, 
			CASE 
				WHEN type = 3 THEN calculationID
				ELSE filterID
			END AS filterID, 
			fieldSelectionFilter
		FROM [dbo].[ASRSysExprComponents]
		WHERE exprID = @piFilterID
			AND ((type = 7) 
				OR ((type = 1) AND (fieldSelectionFilter > 0)) 
				OR (type = 2) 
				OR (type = 3) 
				OR (type = 10))
		ORDER BY componentID;
		
	OPEN components_cursor;
	FETCH NEXT FROM components_cursor INTO @iComponentID, @iType, @iExprID, @iFieldFilterID;
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iType = 1
		BEGIN
			/* Field value with filter. */
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iFieldFilterID, @sComponents OUTPUT;

			IF LEN(@sComponents) > 0
			BEGIN
				SET @psComponents = @psComponents + 
					CASE
						WHEN LEN(@psComponents) > 0 THEN ','
						ELSE ''
					END + 
					@sComponents;
			END
		END

		IF @iType = 7
		BEGIN
			/* Prompted value. */
			SET @psComponents = @psComponents + 
				CASE
					WHEN LEN(@psComponents) > 0 THEN ','
					ELSE ''
				END +
				convert(varchar(255), @iComponentID);
		END

		IF (@iType = 10) OR (@iType = 3)
		BEGIN
			/* Sub-filter or calculation. */
			EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iExprID, @sComponents OUTPUT;

			IF LEN(@sComponents) > 0
			BEGIN
				SET @psComponents = @psComponents + 
					CASE
						WHEN LEN(@psComponents) > 0 THEN ','
						ELSE ''
					END + 
					@sComponents;
			END
		END
	
		IF @iType = 2
		BEGIN
			/* Function. Check if there are any prompted values in the parameter expressions.. */
			DECLARE function_cursor CURSOR LOCAL FAST_FORWARD FOR 
				SELECT exprID 
				FROM [dbo].[ASRSysExpressions]
				WHERE parentComponentID = @iComponentID;
				
			OPEN function_cursor;
			FETCH NEXT FROM function_cursor INTO @iExprID;
			WHILE (@@fetch_status = 0)
			BEGIN
				EXEC [dbo].[sp_ASRIntGetFilterPromptedValues] @iExprID, @sComponents OUTPUT;

				IF LEN(@sComponents) > 0
				BEGIN
					SET @psComponents = @psComponents + 
						CASE
							WHEN LEN(@psComponents) > 0 THEN ','
							ELSE ''
						END + 
						@sComponents;
				END

				FETCH NEXT FROM function_cursor INTO @iExprID;
			END
			CLOSE function_cursor;
			DEALLOCATE function_cursor;
		END
	
		FETCH NEXT FROM components_cursor INTO @iComponentID, @iType, @iExprID, @iFieldFilterID;
	END
	
	CLOSE components_cursor;
	DEALLOCATE components_cursor;

END