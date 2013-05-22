CREATE PROCEDURE [dbo].[sp_ASRFn_InitialsFromForenames]
(
	@psResult 		varchar(MAX) OUTPUT,
	@psForenames	varchar(MAX)
)
AS
BEGIN
	DECLARE @iCounter	integer;

	SET @iCounter = 1

	IF LEN(@psForenames) > 0 
	BEGIN
		SET @psResult = UPPER(left(@psForenames,1));

		WHILE @iCounter < LEN(@psForenames)
		BEGIN
			IF SUBSTRING(@psForenames, @iCounter, 1) = ' '
			BEGIN
				SET @psResult = @psResult + UPPER(SUBSTRING(@psForenames, @iCounter+1, 1));
			END
	
			SET @iCounter = @iCounter + 1;
		END

		SET @psResult = @psResult + ' ';
	END
END
