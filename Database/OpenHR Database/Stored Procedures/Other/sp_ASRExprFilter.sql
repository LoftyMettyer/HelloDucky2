CREATE PROCEDURE sp_ASRExprFilter (
	@resultTableName	varchar(255) OUTPUT,
	@tableID		int=0,
	@filterExprID		int=0)
AS

DECLARE @hResult	integer,
	@filterResult	bit,
	@recordID	integer,
	@tableName	varchar(100),
	@exprName	varchar(100),
	@unfilteredIDs	CURSOR,
	@count		int,
	@idString	varchar(255)

/* Create a unique name for the results. */
SET @count = 1 
SET @resultTableName = 'ASRSys_FilterResultTable_' + convert(varchar(255), @count)
WHILE EXISTS (SELECT * FROM sysobjects WHERE type = 'U' AND name = @resultTableName) 
BEGIN
	SET @count = @count + 1
	SET @resultTableName = 'ASRSys_FilterResultTable_' + convert(varchar(255), @count)
END

/* Create the table that will hold the filtered record IDs. */
EXECUTE ('CREATE TABLE ' + @resultTableName + ' (id INT)')

/* Get the name of the given table. */
SELECT @tableName = tableName
	FROM ASRSysTables
	WHERE tableID = @tableID

/* Get the name of the given filter expression. */
SET @exprName = 'sp_ASRExpr_' + convert(varchar(100), @filterExprID)

IF (NOT @tableName IS NULL)
AND EXISTS (SELECT * FROM sysobjects WHERE type = 'P' AND name = @exprName)
BEGIN
	/* Create a cursor of all records in the given table. */
	EXECUTE ('DECLARE unfilteredIDs CURSOR FAST_FORWARD FOR SELECT id FROM ' + @tableName)
	SET @unfilteredIDs = unfilteredIDs
	DEALLOCATE unfilteredIDs

	OPEN @unfilteredIDs

	/* Loop through the records checking if they satisfy the filter criteria. */
	FETCH NEXT FROM @unfilteredIDs INTO @recordID
	WHILE (@@fetch_status = 0)
	BEGIN
		
		/* Execute the filter expression to see if the current record satisfies it. */
		EXECUTE @exprName @filterResult OUTPUT, @recordID
		IF @hResult <> 0 SELECT @filterResult = 0

		/* Add the record ID to the string of filtered record IDs if the record satisfied the filter criteria. */
		IF @filterResult = 1
		BEGIN
			SET @idString =  CONVERT(varchar(255), @recordID)
			EXECUTE ('INSERT INTO ' + @resultTableName + ' VALUES(' + @idString + ')')
 		END

		FETCH NEXT FROM @unfilteredIDs  INTO @recordID
	END

	/* Free the cursor. */
	CLOSE @unfilteredIDs
	DEALLOCATE @unfilteredIDs
END

/* Return the recordset of filtered IDs. */
EXECUTE ('SELECT id FROM ' + @resultTableName)

















GO

