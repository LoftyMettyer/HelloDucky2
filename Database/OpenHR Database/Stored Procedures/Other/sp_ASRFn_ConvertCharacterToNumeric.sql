CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertCharacterToNumeric]
(
	@pdblResult			float OUTPUT,
	@psStringToConvert  varchar(MAX)
)
AS
BEGIN
	IF (@psStringToConvert IS NULL) OR (LEN(@psStringToConvert) = 0)
	BEGIN
		SET @pdblResult = 0;
	END
	ELSE
	BEGIN
		IF ISNUMERIC(@psStringToConvert) = 1
		BEGIN
			SET @pdblResult = CONVERT(FLOAT, CONVERT(money, @psStringToConvert));
		END
		ELSE
		BEGIN
			SET @pdblResult = 0;
		END
	END
END