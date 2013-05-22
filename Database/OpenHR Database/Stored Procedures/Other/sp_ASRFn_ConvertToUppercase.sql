CREATE PROCEDURE [dbo].[sp_ASRFn_ConvertToUppercase] 
(
	@psResult			varchar(MAX) OUTPUT,
	@psStringToConvert	varchar(MAX)
)
AS
BEGIN
	/* Convert the string to upper case */
	SET @psResult = UPPER(@psStringToConvert);
END
