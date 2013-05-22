CREATE PROCEDURE [dbo].[sp_ASRIntExpressionHasHiddenComponents] (
	@piExprID 			integer, 
	@pfHasHiddenComponents	bit	OUTPUT
)
AS
BEGIN
	/* Check if the given expression has any hidden componeonts. */
	DECLARE @iExprID	integer,
		@sAccess		varchar(MAX),
		@fTemp			bit;

	SET @pfHasHiddenComponents = 0

	DECLARE components_cursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT 
		CASE
			WHEN ASRSysExprComponents.type = 10 THEN	ASRSysExprComponents.filterID
			WHEN ASRSysExprComponents.type = 3 THEN	ASRSysExprComponents.calculationID
			ELSE ASRSysExprComponents.fieldSelectionFilter
		END AS [exprID]
	FROM ASRSysExprComponents
	WHERE exprID = @piExprID
		AND ((type = 3) 
			OR (type = 10) 
			OR ((type = 1) AND (fieldSelectionFilter > 0)))
	OPEN components_cursor
	FETCH NEXT FROM components_cursor INTO @iExprID
	WHILE (@@fetch_status = 0)
	BEGIN
		SELECT @sAccess = access
		FROM ASRSysExpressions
		WHERE exprID = @iExprID	

		IF @sAccess = 'HD'
		BEGIN
			/* The filter/calc is hidden. */
			SET @pfHasHiddenComponents = 1
			RETURN
		END
		ELSE
		BEGIN
			/* The filter/calc is NOT hidden. Check the sub-components. */
			execute sp_ASRIntExpressionHasHiddenComponents @iExprID, @fTemp OUTPUT

			IF @fTemp = 1
			BEGIN
				SET @pfHasHiddenComponents = 1
				RETURN
			END	
		END

		FETCH NEXT FROM components_cursor INTO @iExprID
	END
	CLOSE components_cursor
	DEALLOCATE components_cursor	
END