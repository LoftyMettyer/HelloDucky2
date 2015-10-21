CREATE PROCEDURE [dbo].[sp_ASRRecordAmended]
(
		@piResult integer OUTPUT,	/* Output variable to hold the result. */
		@piTableID integer,			/* TableID being updated. */
		@psRealSource sysname,		/* RealSource being updated. */
		@piID integer,				/* ID the record being updated. */
		@piTimestamp integer		/* Original timestamp of the record being updated. */
)
WITH EXECUTE AS 'dbo'
AS
BEGIN
		/* Check if the given record has been deleted or changed by another user. */
		/* Return 0 if the record has NOT been amended. */
		/* Return 1 if the record has been amended AND is still in the given table/view. */
		/* Return 2 if the record has been amended AND is no longer in the given table/view. */
		/* Return 3 if the record has been deleted from the table. */
		SET NOCOUNT ON;
		DECLARE @iCurrentTimestamp integer,
				@sSQL nvarchar(MAX),
				@psTableName sysname,
				@iCount integer;
		SET @piResult = 0;

		SELECT @psTableName = TableName FROM ASRSysTables WHERE TableID = @piTableID;

		/* Check that the record has not been updated by another user since it was last checked. */
		SET @sSQL = 'SELECT @iCurrentTimestamp = convert(integer, timestamp)' +
						' FROM ' + @psTableName +
						' WHERE id = ' + convert(varchar(MAX), @piID);
		EXECUTE sp_executesql @sSQL, N'@iCurrentTimestamp int OUTPUT', @iCurrentTimestamp OUTPUT;
		
		IF @iCurrentTimestamp IS null
		BEGIN
				/* Record deleted. */
				SET @piResult = 3;
		END
		ELSE
		BEGIN
				IF @iCurrentTimestamp <> @piTimestamp
				BEGIN
						/* Record changed. Check if it is in the given realsource. */
					 SET @sSQL = 'SELECT @piResult = COUNT(id)' +
						 ' FROM ' + @psRealSource +
						 ' WHERE id = ' + convert(varchar(255), @piID);
					 EXECUTE sp_executesql @sSQL, N'@piResult int OUTPUT', @iCount OUTPUT;
					 IF @iCount > 0
					 BEGIN
							 SET @piResult = 1;
					 END
					 ELSE
					 BEGIN
							 SET @piResult = 2;
					 END
				END
		END
END