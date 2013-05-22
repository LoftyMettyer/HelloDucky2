CREATE PROCEDURE [dbo].[sp_ASRFn_CapitalizeInitials]
(
	@psResult	varchar(MAX) OUTPUT,
	@psString	varchar(MAX)
)
AS
BEGIN
	DECLARE @iCounter integer;
	DECLARE @sTemp varchar(1);

	SET @iCounter = 1;

	WHILE @iCounter < LEN(@psString)
	BEGIN
		IF SUBSTRING(@psString, @iCounter, 1) = ' '
		BEGIN
			SET @sTemp = SUBSTRING(@psString, @iCounter+1, 1);
			SET @psString = STUFF(@psString, @iCounter+1, 1, UPPER(@sTemp));
		END
		ELSE
		BEGIN
			SET @sTemp = SUBSTRING(@psString, @iCounter+1, 1);
			SET @psString = STUFF(@psString, @iCounter+1, 1, LOWER(@sTemp));
		END

		SET @iCounter = @iCounter + 1;
	END

	-- Change the first letter too
	SET @sTemp = SUBSTRING(@psString, 1, 1);
	SET @psString = STUFF(@psString, 1, 1, UPPER(@sTemp));

	SET @psResult = @psString;
	
END




GO

