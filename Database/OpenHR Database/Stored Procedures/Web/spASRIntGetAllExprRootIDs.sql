CREATE PROCEDURE spASRIntGetAllExprRootIDs
(
		@iExprID integer,
		@superExpressions cursor varying output
)
AS
BEGIN
	/* Return a cursor of the expressions that use the given expression. */
	DECLARE	@iComponentID	integer,
					@iRootExprID	integer,
					@superCursor	cursor,
					@iTemp				integer

	CREATE TABLE #superExpressionIDs (id integer)

	DECLARE check_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT componentID
		FROM ASRSysExprComponents
		WHERE (calculationID = @iExprID)
			OR (filterID = @iExprID)
			OR ((fieldSelectionFilter = @iExprID) AND (type = 1))

	OPEN check_cursor
	FETCH NEXT FROM check_cursor INTO @iComponentID
	WHILE (@@fetch_status = 0)
	BEGIN
		exec sp_ASRIntGetRootExpressionIDs @iComponentID, @iRootExprID	OUTPUT

		INSERT INTO #superExpressionIDs (id) VALUES (@iRootExprID)
		
		exec spASRIntGetAllExprRootIDs @iRootExprID, @superCursor output
		
		FETCH NEXT FROM @superCursor INTO @iTemp
		WHILE (@@fetch_status = 0)
		BEGIN
			INSERT INTO #superExpressionIDs (id) VALUES (@iTemp)
			
			FETCH NEXT FROM @superCursor INTO @iTemp 
		END
		CLOSE @superCursor
		DEALLOCATE @superCursor

		FETCH NEXT FROM check_cursor INTO @iComponentID
	END
	CLOSE check_cursor
	DEALLOCATE check_cursor
	
	SET @superExpressions = CURSOR FORWARD_ONLY STATIC FOR
		SELECT id FROM #superExpressionIDs
	OPEN @superExpressions
END
GO

