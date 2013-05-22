CREATE PROCEDURE [dbo].[sp_ASRCaseSensitiveCompare]
(
	@pfResult		bit OUTPUT,
	@psStringA 		varchar(MAX),
	@psStringB		varchar(MAX)
)
AS
BEGIN

	-- Return 1 if the given string are exactly equal.
	DECLARE @iPosition	integer;

	SET @pfResult = 0;

	IF (@psStringA IS NULL) AND (@psStringB IS NULL) SET @pfResult = 1;

	IF (@pfResult = 0) AND (NOT @psStringA IS NULL) AND (NOT @psStringB IS NULL)
	BEGIN

		-- LEN() does not look at trailing spaces, so force it too by adding some quotations at the end.
		SET @psStringA = @psStringA + '''';
		SET @psStringB = @psStringB + '''';

		IF LEN(@psStringA) = LEN(@psStringB)
		BEGIN
			SET @pfResult = 1;

			SET @iPosition = 1;
			WHILE @iPosition <= LEN(@psStringA) 
			BEGIN
				IF ASCII(SUBSTRING(@psStringA, @iPosition, 1)) <> ASCII(SUBSTRING(@psStringB, @iPosition, 1))
				BEGIN
					SET @pfResult = 0;
					BREAK
				END

				SET @iPosition = @iPosition + 1;
			END
		END
	END
END