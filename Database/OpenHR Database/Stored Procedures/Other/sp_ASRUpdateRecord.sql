CREATE PROCEDURE [dbo].[sp_ASRUpdateRecord]
(
    @piResult integer OUTPUT,		/* Output variable to hold the result. */
    @psUpdateString nvarchar(MAX),  /* SQL Update string to update the record. */
    @piTableID integer,				/* TableID being updated. */
    @psRealSource sysname,			/* RealSource being updated. */
    @piID integer,					/* ID the record being updated. */
    @piTimestamp integer			/* Original timestamp of the record being updated. */
)
AS
BEGIN
    SET NOCOUNT ON;

	-- Get status of amended record
	EXEC dbo.sp_ASRRecordAmended @piResult OUTPUT,
	    @piTableID,
		@psRealSource,
		@piID,
		@piTimestamp;

    -- Run the given SQL UPDATE string.   
    IF @piResult = 0
		EXECUTE sp_executeSQL @psUpdateString;

END