CREATE PROCEDURE [dbo].spASRColumnsUsedInExpression
(
	@piExpressionID		integer,
	@pcurColumns		cursor varying output
)
AS
BEGIN
	-- Return the IDs of the columns used in the given expression.
	DECLARE
		@iType			integer,
		@iID			integer,
		@iExprID		integer,
		@iColumnID		integer,
		@curSubColumns	cursor

	CREATE TABLE #curColumns (columnID integer)

	-- Record the columns used by field components.
	INSERT INTO #curColumns
	SELECT DISTINCT EC.fieldColumnID
	FROM ASRSysExprComponents EC
	WHERE EC.exprID = @piExpressionID
		AND EC.type = 1 -- Field component

	-- Check sub-expressions.
	DECLARE curSubExpressions CURSOR LOCAL FAST_FORWARD FOR 
	SELECT
		EC.type, 
		CASE 
			WHEN EC.type = 1 THEN EC.fieldSelectionFilter -- Field filter
			WHEN EC.type = 2 THEN EC.componentID -- Function
			WHEN EC.type = 3 THEN EC.calculationID -- Calculation
			WHEN EC.type = 10 THEN EC.filterID -- Filter
		END
	FROM ASRSysExprComponents EC
	WHERE EC.exprID = @piExpressionID
		AND ((EC.type = 1 AND EC.fieldSelectionFilter > 0)
			OR (EC.type = 2)
			OR (EC.type = 3)
			OR (EC.type = 10))

	OPEN curSubExpressions
	FETCH NEXT FROM curSubExpressions INTO @iType, @iID
	WHILE (@@fetch_status = 0)
	BEGIN
		IF @iType = 2
		BEGIN
			-- Get the columns used in as follows:
			-- 1) Function component sub-expressions
			DECLARE curFunctionSubExpressions CURSOR LOCAL FAST_FORWARD FOR 
			SELECT
				E.exprID
			FROM ASRSysExpressions E
			WHERE E.parentComponentID = @iID

			OPEN curFunctionSubExpressions
			FETCH NEXT FROM curFunctionSubExpressions INTO @iExprID
			WHILE (@@fetch_status = 0)
			BEGIN
				EXEC spASRColumnsUsedInExpression @iExprID, @curSubColumns OUTPUT

				FETCH NEXT FROM @curSubColumns INTO @iColumnID
				WHILE (@@fetch_status = 0)
				BEGIN
					INSERT INTO #curColumns (columnID) VALUES (@iColumnID)
							
					FETCH NEXT FROM @curSubColumns INTO @iColumnID
				END
				CLOSE @curSubColumns
				DEALLOCATE @curSubColumns

				FETCH NEXT FROM curFunctionSubExpressions INTO @iExprID
			END
			CLOSE curFunctionSubExpressions
			DEALLOCATE curFunctionSubExpressions
		END
		ELSE
		BEGIN
			-- Get the columns used in as follows:
			-- 1) Field component filters
			-- 2) Calculation components
			-- 3) Filter components
			EXEC spASRColumnsUsedInExpression @iID, @curSubColumns OUTPUT

			FETCH NEXT FROM @curSubColumns INTO @iColumnID
			WHILE (@@fetch_status = 0)
			BEGIN
				INSERT INTO #curColumns (columnID) VALUES (@iColumnID)
						
				FETCH NEXT FROM @curSubColumns INTO @iColumnID
			END
			CLOSE @curSubColumns
			DEALLOCATE @curSubColumns
		END

		FETCH NEXT FROM curSubExpressions INTO @iType, @iID
	END
	CLOSE curSubExpressions
	DEALLOCATE curSubExpressions

	/* Return the cursor of columns. */
	SET @pcurColumns = CURSOR FORWARD_ONLY STATIC FOR
		SELECT columnID 
		FROM #curColumns
	OPEN @pcurColumns

	DROP TABLE #curColumns
END

GO

