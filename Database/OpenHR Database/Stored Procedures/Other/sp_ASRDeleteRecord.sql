CREATE PROCEDURE [dbo].[sp_ASRDeleteRecord]
(
    @piResult integer OUTPUT,   /* Output variable to hold the result. */
    @piTableID integer,			/* TableID being deleted from. */
    @psRealSource sysname,		/* RealSource being deleted from. */
    @piID integer				/* ID the record being deleted. */
)
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @iTimestamp integer,
			@sSQL		nvarchar(MAX);

	-- Get status of amended record
	EXEC dbo.sp_ASRRecordAmended @piResult OUTPUT,
	    @piTableID,
		@psRealSource,
		@piID,
		@iTimestamp;

	-- If Ok run the delete statement
    IF @piResult <> 3
    BEGIN
       SET @sSQL = 'DELETE ' +
            ' FROM ' + @psRealSource +
            ' WHERE id = ' + convert(varchar(MAX), @piID);
       EXECUTE sp_executesql @sSQL;
    END

END