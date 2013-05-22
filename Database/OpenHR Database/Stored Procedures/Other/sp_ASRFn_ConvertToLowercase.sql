
CREATE PROCEDURE sp_ASRFn_ConvertToLowercase 
(
	@psResult			varchar(MAX) OUTPUT,
	@psStringToConvert 	varchar(MAX)
)
AS
BEGIN
	SET @psResult = LOWER(@psStringToConvert);
END
