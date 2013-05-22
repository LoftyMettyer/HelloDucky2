CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertToPropercase]
(
	@psOutput	varchar(MAX) OUTPUT,
	@psInput 	varchar(MAX)
)
AS
BEGIN

	DECLARE @Index	integer,
			@Char	char(1);

	SET @psOutput = LOWER(@psInput);
	SET @Index = 1;
	SET @psOutput = STUFF(@psOutput, 1, 1,UPPER(SUBSTRING(@psInput,1,1)));

	WHILE @Index <= LEN(@psInput)
	BEGIN

		SET @Char = SUBSTRING(@psInput, @Index, 1);

		IF @Char IN ('m','M',' ', ';', ':', '!', '?', ',', '.', '_', '-', '/', '&','''','(',char(9), char(13), char(10))
		BEGIN
			IF @Index + 1 <= LEN(@psInput)
			BEGIN
				IF @Char = '' AND UPPER(SUBSTRING(@psInput, @Index + 1, 1)) != 'S'
					SET @psOutput = STUFF(@psOutput, @Index + 1, 1,UPPER(SUBSTRING(@psInput, @Index + 1, 1)));
				ELSE IF UPPER(@Char) != 'M'
					SET @psOutput = STUFF(@psOutput, @Index + 1, 1,UPPER(SUBSTRING(@psInput, @Index + 1, 1)));

				-- Catch the McName
				IF UPPER(@Char) = 'M' AND UPPER(SUBSTRING(@psInput, @Index + 1, 1)) = 'C' AND UPPER(SUBSTRING(@psInput, @Index - 1, 1)) = ''
				BEGIN
					SET @psOutput = STUFF(@psOutput, @Index + 1, 1,LOWER(SUBSTRING(@psInput, @Index + 1, 1)));
					SET @psOutput = STUFF(@psOutput, @Index + 2, 1,UPPER(SUBSTRING(@psInput, @Index + 2, 1)));
					SET @Index = @Index + 1;
				END
			END
		END

	SET @Index = @Index + 1;
	END

END