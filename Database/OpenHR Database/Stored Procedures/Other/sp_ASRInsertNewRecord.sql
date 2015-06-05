CREATE PROCEDURE [dbo].[sp_ASRInsertNewRecord]
	(
		@piNewRecordID integer OUTPUT,   /* Output variable to hold the new record ID. */
		@psInsertString nvarchar(MAX)    /* SQL Insert string to insert the new record. */
	)
	AS
	BEGIN
		SET NOCOUNT ON;

		-- Run the given SQL INSERT
		EXECUTE sp_executesql @psInsertString;

		-- Calculate the ID
		SELECT @piNewRecordID = convert(int,convert(varbinary(4),CONTEXT_INFO()));

END

