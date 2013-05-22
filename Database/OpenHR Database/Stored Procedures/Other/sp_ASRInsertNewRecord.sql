CREATE PROCEDURE [dbo].[sp_ASRInsertNewRecord]
	(
		@piNewRecordID integer OUTPUT,   /* Output variable to hold the new record ID. */
		@psInsertString nvarchar(MAX)    /* SQL Insert string to insert the new record. */
	)
	AS
	BEGIN
		SET NOCOUNT ON;

		DECLARE @ssql nvarchar(MAX),
				@tablename varchar(255);

		-- Run the given SQL INSERT
		EXECUTE sp_executesql @psInsertString;

		-- Calculate the ID
		SET  @psInsertString = REPLACE(' ' + @psInsertString,' INSERT INTO ','')
		SET  @psInsertString = REPLACE(' ' + @psInsertString,' INSERT ','')
		SET @tablename = SUBSTRING(@psInsertString,0, CHARINDEX('(', @psInsertString));

		IF NOT @tablename = ''
		BEGIN
			SET @ssql = 'SELECT @ID = MAX(ID) FROM ' + @tablename;
			EXECUTE sp_executesql @ssql, N'@ID int OUTPUT', @ID = @piNewRecordID OUTPUT;
		END
		ELSE SET @piNewRecordID = 0	

END

