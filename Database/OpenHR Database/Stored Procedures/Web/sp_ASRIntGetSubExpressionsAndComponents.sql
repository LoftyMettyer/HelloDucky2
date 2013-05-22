CREATE PROCEDURE [dbo].[sp_ASRIntGetSubExpressionsAndComponents] (
	@piExprID				integer, 
	@psTempExprIDs 			varchar(MAX)	OUTPUT, 
	@psTempComponentIDs		varchar(MAX)	OUTPUT
)
AS
BEGIN
	DECLARE
		@iComponentID		integer,
		@iExpressionID		integer,
		@sSubExprIDs 		varchar(MAX), 
		@sSubComponentIDs	varchar(MAX);	

	SET @psTempExprIDs = '';
	SET @psTempComponentIDs = '';

	DECLARE components_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysExprComponents.componentID
		FROM ASRSysExprComponents
		WHERE ASRSysExprComponents.exprID = @piExprID;
	OPEN components_cursor;
	FETCH NEXT FROM components_cursor INTO @iComponentID;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @psTempComponentIDs = @psTempComponentIDs +
			CASE
				WHEN len(@psTempComponentIDs) > 0 THEN ','
				ELSE ''
			END +
			convert(varchar(100), @iComponentID);
			
		FETCH NEXT FROM components_cursor INTO @iComponentID;
	END
	CLOSE components_cursor;
	DEALLOCATE components_cursor;

	DECLARE expressions_cursor CURSOR LOCAL FAST_FORWARD FOR 
		SELECT ASRSysExpressions.exprID
		FROM ASRSysExpressions
		WHERE ASRSysExpressions.parentComponentID IN
			(SELECT ASRSysExprComponents.componentID
			FROM ASRSysExprComponents
			WHERE ASRSysExprComponents.exprID = @piExprID);
	OPEN expressions_cursor;
	FETCH NEXT FROM expressions_cursor INTO @iExpressionID;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @psTempExprIDs = @psTempExprIDs +
			CASE
				WHEN len(@psTempExprIDs) > 0 THEN ','
				ELSE ''
			END +
			convert(varchar(100), @iExpressionID);
		
		exec [dbo].[sp_ASRIntGetSubExpressionsAndComponents] @iExpressionID, @sSubExprIDs OUTPUT, @sSubComponentIDs OUTPUT;
		
		IF len(@sSubExprIDs) > 0
		BEGIN
			SET @psTempExprIDs = @psTempExprIDs +
				CASE
					WHEN len(@psTempExprIDs) > 0 THEN ','
					ELSE ''
				END +
				@sSubExprIDs;
		END

		IF len(@sSubComponentIDs) > 0
		BEGIN
			SET @psTempComponentIDs = @psTempComponentIDs +
				CASE
					WHEN len(@psTempComponentIDs) > 0 THEN ','
					ELSE ''
				END +
				@sSubComponentIDs;
		END

		FETCH NEXT FROM expressions_cursor INTO @iExpressionID;
	END
	CLOSE expressions_cursor;
	DEALLOCATE expressions_cursor;
END