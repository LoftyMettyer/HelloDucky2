CREATE PROCEDURE [dbo].[spASRIntGetLookupFindColumnInfo] (
	@piLookupColumnID 		integer,
	@ps1000SeparatorCols	varchar(MAX)	OUTPUT
)
AS
BEGIN
	DECLARE 
		@iTableID			integer,
		@bUse1000Separator	bit,
		@iOrderID			integer;
		
	/* Get the column name. */
	SELECT @iTableID = tableID,
		@bUse1000Separator = Use1000Separator
	FROM [dbo].[ASRSysColumns]
	WHERE columnID = @piLookupColumnID;

	SET @ps1000SeparatorCols = 
		CASE
			WHEN @bUse1000Separator = 1 THEN '1'
			ELSE '0'
		END;

	/* Get the table name and default order. */
	SELECT @iOrderID = defaultOrderID
	FROM [dbo].[ASRSysTables]
	WHERE tableID = @iTableID;

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT ASRSysColumns.Use1000Separator
	FROM ASRSysOrderItems
	INNER JOIN ASRSysColumns ON ASRSysOrderItems.columnID = ASRSysColumns.columnID
	INNER JOIN ASRSysTables ON ASRSysTables.tableID = ASRSysColumns.tableID
	WHERE ASRSysOrderItems.orderID = @iOrderID
		AND ASRSysOrderItems.type = 'F'
		AND ASRSysOrderItems.columnID <> @piLookupColumnID
	ORDER BY ASRSysOrderItems.sequence;

	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @bUse1000Separator;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
			CASE
				WHEN @bUse1000Separator = 1 THEN '1'
				ELSE '0'
			END;

		FETCH NEXT FROM orderCursor INTO @bUse1000Separator;
	END
	CLOSE orderCursor;
	DEALLOCATE orderCursor;

END