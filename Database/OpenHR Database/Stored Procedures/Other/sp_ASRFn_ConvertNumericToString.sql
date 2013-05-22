CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertNumericToString]
(
	@psResult				varchar(MAX) OUTPUT,
    @pdblNumericToConvert	float,
   	@piDecimalPlaces 		integer
)
AS
BEGIN
	/* Convert the number to a string */
	SET @psResult = LTRIM(STR(@pdblNumericToConvert, 20, @piDecimalPlaces));
END