CREATE PROCEDURE [dbo].[spASRIntGetLookupFindColumnInfo] (
	@piLookupColumnID 		integer,
	@ps1000SeparatorCols	varchar(MAX)	OUTPUT,
	@psBlanIfZeroCols		varchar(MAX)	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE 
		@iTableID			integer,
		@bUse1000Separator	bit,
		@bBlankIfZero		bit,
		@iOrderID			integer;
		
	/* Get the column name. */
	SELECT @iTableID = tableID,
		@bUse1000Separator = Use1000Separator,
		@bBlankIfZero = BlankIfZero
	FROM [dbo].[ASRSysColumns]
	WHERE columnID = @piLookupColumnID;

	SET @ps1000SeparatorCols = 
		CASE
			WHEN @bUse1000Separator = 1 THEN '1'
			ELSE '0'
		END;

	SET @psBlanIfZeroCols =
		CASE
			WHEN @bBlankIfZero = 1 THEN '1'
			ELSE '0'
		END;

	/* Get the table name and default order. */
	SELECT @iOrderID = defaultOrderID
	FROM [dbo].[ASRSysTables]
	WHERE tableID = @iTableID;

	/* Create the order select strings. */
	DECLARE orderCursor CURSOR LOCAL FAST_FORWARD FOR 
	SELECT c.Use1000Separator, c.BlankIfZero
	FROM ASRSysOrderItems oi
		INNER JOIN ASRSysColumns c ON oi.columnID = c.columnId
		INNER JOIN ASRSysTables t ON t.tableID = c.tableID
	WHERE oi.orderID = @iOrderID AND oi.type = 'F'
		AND oi.columnID <> @piLookupColumnID
		AND c.dataType <> -4 AND c.datatype <> -3
	ORDER BY oi.sequence;

	OPEN orderCursor;
	FETCH NEXT FROM orderCursor INTO @bUse1000Separator, @bBlankIfZero;
	WHILE (@@fetch_status = 0)
	BEGIN
		SET @ps1000SeparatorCols = @ps1000SeparatorCols + 
			CASE
				WHEN @bUse1000Separator = 1 THEN '1'
				ELSE '0'
			END;

		SET @psBlanIfZeroCols = @psBlanIfZeroCols +
			CASE
				WHEN @bBlankIfZero = 1 THEN '1'
				ELSE '0'
			END;

		FETCH NEXT FROM orderCursor INTO @bUse1000Separator, @bBlankIfZero;
	END
	CLOSE orderCursor;
	DEALLOCATE orderCursor;

END